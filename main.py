import os
import shutil
import base64
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
from PIL import Image
from PIL.ExifTags import TAGS
import pytz
import ffmpeg
import exiftool
import asyncio
from concurrent.futures import ThreadPoolExecutor
import platform
import multiprocessing
import filecmp
import locale
import re

# --- 各種設定 ---
# 画像・動画の拡張子リスト
IMAGE_EXTS = ['.jpg', '.jpeg', '.png', '.tif', '.tiff', '.heic', '.dng', '.arw']
VIDEO_EXTS = ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.webm', '.mts', '.mpg']
METADATA_EXTS = ['.xml', '.thm']


# グローバルでディレクトリごとの asyncio.Lock を管理する辞書
dir_locks = {}
dir_locks_lock = asyncio.Lock() # この辞書へのアクセス用ロック

# --- 同期処理(ブロッキング関数群) ---
def decode_value(value):
    """安全にEXIF値をデコードする"""
    if value is None:
        return None
    # バイト列の処理
    if isinstance(value, bytes):
        # 印字可能な文字列かどうかをチェック
        try:
            # UTF-8, UTF-16, ASCII, Latin-1で順次試行
            decode_attempts = [
                ("utf-8", "strict"),
                ("shift_jis", "replace"),
                ("utf-16le", "replace"),
                ("utf-16be", "replace"),
                ("ascii", "replace"),
                ("latin-1", "replace"),
                ("mac-roman", "replace")
            ]
            # 複数のエンコーディングで試行
            for encoding, errors in decode_attempts:
                try:
                    decoded = value.decode(encoding, errors=errors)
                    # 印字可能な文字のみを含む場合に返す
                    # 印字可能チェックを少し緩和(制御文字などを許容)
                    if all(c.isprintable() or c.isspace() for c in decoded.strip()):
                        # 先頭・末尾のヌル文字や空白を除去
                        return decoded.strip().strip('\x00')
                except UnicodeDecodeError:
                    continue
                except Exception:
                    continue
            # デコードできない場合は、Base64エンコードして返す
            return base64.b64encode(value).decode("ascii")
        except:
            # 最終的にデコードできない場合
            return None
            # return f"[Undecoded Data: {len(value)} bytes]"
    # 他の型の値をそのまま返す
    return str(value)

def get_image_date(file_path):
    """
    PILまたはExifToolを使い、画像ファイルから最も古い有効な撮影日時を取得。
    成功時は 'YYYY_MM_DD_HH_MM_SS' 形式の文字列を返す。
    取得できなければ、Noneを返す。
    """
    valid_datetimes = [] # 有効な日時(datetimeオブジェクト)を格納するリスト
    print(f"画像 {os.path.basename(file_path)}: 日時情報収集開始...")
    # 1. PILで日時タグを検索し、リストに追加
    try:
        with Image.open(file_path) as im:
            exif = im.getexif()
            if exif:
                # DateTimeOriginal (0x9003), DateTimeDigitized (0x9004), DateTime (0x0132) のタグ情報で試す。
                # これらは撮影日時やデジタル化日時を示唆するため候補とする。
                pil_tag_ids = [0x9003, 0x9004, 0x0132]
                for tag_id in pil_tag_ids:
                    if tag_id in exif:
                        datetime_raw = exif.get(tag_id)
                        if datetime_raw:
                            datetime_str = decode_value(datetime_raw)
                            if datetime_str:
                                pil_dt = validate_and_parse_datetime(datetime_str)
                                if pil_dt:
                                    print(f" [PIL] Tag {hex(tag_id)} -> 候補: {pil_dt.strftime('%Y:%m:%d %H:%M:%S')}")
                                    valid_datetimes.append(pil_dt)
    except FileNotFoundError:
        print(f"ファイルが見つかりません [PIL]: {file_path}")
        return None
    except Image.UnidentifiedImageError:
        print(f"認識できない画像形式 [PIL]: {os.path.basename(file_path)}")
    except Exception as e:
        # OSError: broken data stream などPILが扱えない場合でもログは出す。
        print(f"EXIF取得エラー [PIL]: {os.path.basename(file_path)} ({e})")
    
    # 2. ExifToolで日時タグを検索し、リストに追加
    if not valid_datetimes:
        try:
            files = [str(file_path)]
            with exiftool.ExifToolHelper() as et:
                # 主要な作成日時系のタグを指定(-FileModifyDateは含めない)
                # ModifyDateも、他に何もなければ候補になりうるが、今回は除外
                params = [
                    "-DateTimeOriginal", "-CreateDate", "-DateCreated", # 標準的な作成日時タグ
                    "-SubSecDateTimeOriginal", "-SubSecCreateDate", # サブ秒含むタグ
                    "-MakerNotes:DateTimeOriginal", # メーカー独自タグも試す
                    "-fast", "-api", "largefilesupport=1"
                ]
                metadata = et.get_metadata(files, params=params)
            if metadata:
                d = metadata[0]
                # 収集対象とするExifToolのタグ名リスト
                # FileMofidyDate は最も古い日時を求める意図から除外
                exiftool_tags_to_check = [
                    "EXIF:DateTimeOriginal", "EXIF:CreateDate",
                    "XMP:DateTimeOriginal", "XMP:CreateDate", "XMP:DateCreated",
                    "MakerNotes:DateTimeOriginal",
                    "Composite:SubSecDateTimeOriginal", "Composite:SubSecCreateDate",
                ]
                for key in exiftool_tags_to_check:
                    tag_name_only = key.split(':')[-1]
                    if tag_name_only in d:
                        val = d[tag_name_only]
                        dt_candidate = None
                        if isinstance(val, str) and val.strip():
                            dt_candidate = validate_and_parse_datetime(val)
                        elif isinstance(val, datetime):
                            if val.tzinfo:
                                tokyo_tz = pytz.timezone('Asia/Tokyo')
                                val = val.astimezone(tokyo_tz).replace(tzinfo=None)
                            if validate_and_parse_datetime(val.strftime("%Y:%m:%d %H:%M:%S")):
                                dt_candidate = val
                        if dt_candidate:
                            print(f" [ExifTool] Tag {key} -> 候補: {dt_candidate.strftime('%Y:%m:%d %H:%M:%S')}")
                            valid_datetimes.append(dt_candidate)
            else:
                print(f"メタデータ取得エラー [ExifTool]: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"メタデータ取得中に予期せぬエラー [ExifTool]: {os.path.basename(file_path)} ({e})")
    
    # 3. 収集した有効な日時の中から最も古いものを選択
    if valid_datetimes:
        # 重複を除去してソートし、最小値(最も古い日時)を取得
        oldest_datetime = min(list(set(valid_datetimes)))
        print(f"-> 画像 {os.path.basename(file_path)}: 最も古い日時 {oldest_datetime.strftime('%Y:%m:%d %H:%M:%S')} を採用")
        # 最も古いdatetimeオブジェクトを期待する文字列形式に変換して返す
        return oldest_datetime.strftime('%Y_%m_%d_%H_%M_%S')
    # ––– メタデータ日時が見つからなかった場合の処理 –––
    if not valid_datetimes:
        print(f"有効なメタデータ日時が見つかりませんでした。ExifToolでファイルのタイムスタンプを確認します: {os.path.basename}")
        file_timestamps_from_exif = []
        try:
            files = [str(file_path)]
            with exiftool.ExifToolHelper() as et:
                # Fileグループのタイムスタンプのみを取得
                params = [
                    "-G", # グループ名を取得
                    "-File:FileModifyDate",
                    "-File:FileAccessDate",
                    "-File:FileInodeChangeDate", # Linux/macOSでのinode変更日時
                    "-File:FileCreateDate",    # Windowsでの作成日時 (ExifToolバージョン依存)
                    "-fast", "-api", "largefilesupport=1"
                ]
                metadata = et.get_metadata(files, params=params)
            if metadata:
                d = metadata[0]
                file_tags_to_check = [
                "File:FileModifyDate",
                "File:FileAccessDate",
                "File:FileInodeChangeDate",
                "File:FileCreateDate",
                ]
                for tag in file_tags_to_check:
                    if tag in d:
                        val = d[tag]
                        # タイムゾーン対応のパース関数を使用
                        dt_candidate = validate_and_parse_datetime(val)
                        if dt_candidate:
                            print(f" [ExifTool File] {tag} -> 候補: {dt_candidate.strftime('%Y:%m:%d %H:%M:%S')}")
                            file_timestamps_from_exif.append(dt_candidate)
            else:
                print(f"ファイルタイムスタンプ取得結果なし [ExifTool File]:{os.path.basename(file_path)}")
        except FileNotFoundError:
            print(f"ファイルが見つかりません [ExifTool File]: {file_path}")
            return None
        except Exception as e:
            print(f"ファイルタイムスタンプ取得中に予期せぬエラー [ExifTool File]: {os.path.basename(file_path)}")
            return None
        if file_timestamps_from_exif:
            unique_file_timestamps = list(set(file_timestamps_from_exif))
            oldest_file_time = min(unique_file_timestamps)
            print(f"-> 動画 {os.path.basename(file_path)}: ExifToolのファイルタイムスタンプから最も古い日時 {oldest_file_time.strftime('%Y:%m:%d %H:%M:%S')} を代替として採用します。")
            return oldest_file_time.strftime('%Y_%m_%d_%H_%M_%S')
        else:
            print(f"有効な日時データを取得できませんでした: {os.path.basename(file_path)}")
            return None

def get_video_date(file_path, local_timezone='Asia/Tokyo'):
    """
    ffmpegまたはExifToolを使用し、動画ファイルから最も古い有効な撮影日時を取得。
    成功時は 'YYYY_MM_DD_HH_MM_SS' 形式の文字列を返す。
    取得できない場合は、ファイルのタイムスタンプ(更新日時、アクセス日時、作成/inode変更日時)の中で最も古いものを代替として使用する。
    """
    valid_datetimes = [] # 有効な日時(datetimeオブジェクト)を格納するリスト
    print(f"動画 {os.path.basename(file_path)}: 日時情報収集開始...")
    # 1. ffmpegで日時情報を収集し、リストに追加(MP4の場合)
    print(f" [ffmpeg] 日時情報を検索...")
    try:
        probe = ffmpeg.probe(file_path)
        format_info = probe.get('format', {})
        creation_time_str = format_info.get('tags', {}).get('creation_time')
        if creation_time_str:
            ffmpeg_dt = validate_and_parse_datetime(creation_time_str)
            if ffmpeg_dt:
                print(f" [ffmpeg] format:tags:creatiem_time -> 候補: {ffmpeg_dt.strftime('%Y:%m:%d %H:%M:%S')}")
                valid_datetimes.append(ffmpeg_dt)
        # stream タグの creation_time(動画/音声トラックにある場合)
        for stream in probe.get('streams', []):
            stream_creation_time = stream.get('tags', {}).get('creation_time')
            if stream_creation_time:
                stream_dt = validate_and_parse_datetime(stream_creation_time)
                if stream_dt:
                    print(f" [ffmpeg] stream[{stream.get('index', '?')}]:tags:creation_time -> 候補: {stream_dt.strftime('%Y:%m:%d %H:%M:%S')}")
                    valid_datetimes.append(stream_dt)
    except ffmpeg.Error as e:
        print(f"ffmpeg probeエラー: {os.path.basename(file_path)} ({e})")
    except Exception as e:
        print(f"ffmpeg処理中の予期せぬエラー: {os.path.basename(file_path)} ({e})")
    # 2. ExifToolを使用する
    if not valid_datetimes:
        try:
            files = [str(file_path)]
            with exiftool.ExifToolHelper() as et:
                # 動画関連の主要な作成日時系のタグを指定 (-FileModifyDateは含めない)
                params = [
                    "-QuickTime:CreateDate", "-QuickTime:MediaCreateDate", "-QuickTime:TrackCreateDate",
                    "-Keys:CreationDate", "-UserData:DateTimeOriginal",
                    "-XMP:DateTimeOriginal", "-XMP:CreateDate", "-XMP:DateCreated",
                    "-H264:DateTimeOriginal", "-MPEG:DateTimeOriginal",
                    "-RIFF:DateTimeOriginal", "-ASF:CreationDate", "-Matroska:DateUTC",
                    "-EXIF:DateTimeOriginal", "-EXIF:CreateDate", # 動画にEXIFがある場合
                    "-Composite:SubSecCreateDate", "-Composite:SubSecDateTimeOriginal",
                    "-fast", "-api", "largefilesupport=1"
                ]
                metadata = et.get_metadata(files, params=params)
            if metadata:
                d = metadata[0]
                # 収集対象とするExifToolのタグ名リスト
                # FileModifyDate は最も古い日時を求める意図から除外
                exiftool_tags_to_check = [
                    "QuickTime:CreateDate", "QuickTime:MediaCreateDate", "QuickTime:TrackCreateDate",
                    "Keys:CreationDate", "UserData:DateTimeOriginal",
                    "XMP:DateTimeOriginal", "XMP:CreateDate", "XMP:DateCreated",
                    "H264:DateTimeOriginal", "MPEG:DateTimeOriginal",
                    "RIFF:DateTimeOriginal", "ASF:CreationDate", "Matroska:DateUTC",
                    "EXIF:DateTimeOriginal", "EXIF:CreateDate",
                    "Composite:SubSecCreateDate", "Composite:SubSecDateTimeOriginal",
                ]
                for key in exiftool_tags_to_check:
                    tag_name_only = key.split(':')[-1]
                    if tag_name_only in d:
                        val = d[tag_name_only]
                        dt_candidate = None
                        if isinstance(val, str) and val.strip():
                            dt_candidate = validate_and_parse_datetime(val)
                        elif isinstance(val, datetime):
                            if val.tzinfo:
                                tz = pytz.timezone(local_timezone)
                                val = val.astimezone(tz).replace(tzinfo=None)
                            if validate_and_parse_datetime(val.strftime("%Y:%m:%d %H:%M:%S")):
                                dt_candidate = val
                        if dt_candidate:
                            print(f" [ExifTool] Tag {key} -> 候補: {dt_candidate.strftime('%Y:%m:%d %H:%M:%S')}")
                            valid_datetimes.append(dt_candidate)
            else:
                print(f"メタデータ取得エラー [ExifTool]: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"メタデータ取得中に予期せぬエラー [ExifTool]: {os.path.basename(file_path)} ({e})")
    # 3. 収集した有効な日時の中から最も古いものを選択または代替処理
    if valid_datetimes:
        # 重複を除去して最小値(最も古い日時)を取得
        unique_valid_datetimes = list(set(valid_datetimes))
        oldest_datetime = min(unique_valid_datetimes)
        print(f"-> 動画 {os.path.dirname(file_path)}: 最も古い日時 {oldest_datetime.strftime('%Y:%m:%d %H:%M:%S')} を採用")
        # 最も古いdatetimeオブジェクトを期待する文字列形式に変換して返す
        return oldest_datetime.strftime('%Y_%m_%d_%H_%M_%S')
    # ––– メタデータ日時が見つからなかった場合の処理 –––
    if not valid_datetimes:
        print(f"有効なメタデータ日時が見つかりませんでした。ExifToolでファイルのタイムスタンプを確認します: {os.path.basename}")
        file_timestamps_from_exif = []
        try:
            files = [str(file_path)]
            with exiftool.ExifToolHelper() as et:
                # Fileグループのタイムスタンプのみを取得
                params = [
                    "-G", # グループ名を取得
                    "-File:FileModifyDate",
                    "-File:FileAccessDate",
                    "-File:FileInodeChangeDate", # Linux/macOSでのinode変更日時
                    "-File:FileCreateDate",    # Windowsでの作成日時 (ExifToolバージョン依存)
                    "-fast", "-api", "largefilesupport=1"
                ]
                metadata = et.get_metadata(files, params=params)
            if metadata:
                d = metadata[0]
                file_tags_to_check = [
                "File:FileModifyDate",
                "File:FileAccessDate",
                "File:FileInodeChangeDate",
                "File:FileCreateDate",
                ]
                for tag in file_tags_to_check:
                    if tag in d:
                        val = d[tag]
                        # タイムゾーン対応のパース関数を使用
                        dt_candidate = validate_and_parse_datetime(val)
                        if dt_candidate:
                            print(f" [ExifTool File] {tag} -> 候補: {dt_candidate.strftime('%Y:%m:%d %H:%M:%S')}")
                            file_timestamps_from_exif.append(dt_candidate)
            else:
                print(f"ファイルタイムスタンプ取得結果なし [ExifTool File]:{os.path.basename(file_path)}")
        except FileNotFoundError:
            print(f"ファイルが見つかりません [ExifTool File]: {file_path}")
            return None
        except Exception as e:
            print(f"ファイルタイムスタンプ取得中に予期せぬエラー [ExifTool File]: {os.path.basename(file_path)}")
            return None
        if file_timestamps_from_exif:
            unique_file_timestamps = list(set(file_timestamps_from_exif))
            oldest_file_time = min(unique_file_timestamps)
            print(f"-> 動画 {os.path.basename(file_path)}: ExifToolのファイルタイムスタンプから最も古い日時 {oldest_file_time.strftime('%Y:%m:%d %H:%M:%S')} を代替として採用します。")
            return oldest_file_time.strftime('%Y_%m_%d_%H_%M_%S')
        else:
            print(f"有効な日時データを取得できませんでした: {os.path.basename(file_path)}")
            return None

def get_file_date(file_path):
    """
    画像または動画ファイルから撮影日時を取得。
    取得できなければ、ファイルの最終更新日時を利用する。
    撮影日時は 'YYYY_MM_DD_HH_MM_SS' 形式で返す。
    """
    ext = os.path.splitext(file_path)[1].lower()
    date_str = None
    if ext in IMAGE_EXTS:
        date_str = get_image_date(file_path)
    elif ext in VIDEO_EXTS:
        date_str = get_video_date(file_path)
    if date_str:
        print(f"-> 取得日時: {date_str} ({os.path.basename(file_path)})")
        return date_str
    else:
        # 最終更新日時を使う場合（オプション）
        # try:
        #     mod_time = os.path.getmtime(file_path)
        #     dt_mod = datetime.fromtimestamp(mod_time)
        #     print(f"-> 最終更新日時を使用: {dt_mod.strftime('%Y_%m_%d_%H_%M_%S')} ({os.path.basename(file_path)})")
        #     return dt_mod.strftime('%Y_%m_%d_%H_%M_%S')
        # except Exception as e:
        #     print(f"-> 最終更新日時の取得エラー: {e} ({os.path.basename(file_path)})")
        print(f"-> 日時取得失敗 ({os.path.basename(file_path)})")
        print(f"日時取得失敗 ({os.path.basename(file_path)})")
        return None

def make_destination_path(dest_root, date_str):
    """
    日付文字列（例: "2019_08_26_09_54_50"）から、dest_root/year/month/day/ を作成し、返す。
    """
    try:
        dt = datetime.strptime(date_str, '%Y_%m_%d_%H_%M_%S')
        year = dt.strftime('%Y')
        month = dt.strftime('%m')
        day = dt.strftime('%d')
        new_basename = dt.strftime('%Y%m%d_%H%M%S')
        dest_dir = os.path.join(dest_root, f"{year}-{month}", day)
        os.makedirs(dest_dir, exist_ok=True)
        return dest_dir, new_basename
    except ValueError: # パース失敗
        print(f"日付文字列 '{date_str}' のパースエラー。移動先パス作成失敗。")
        return None
    except Exception as e:
        print(f"移動先パス作成エラー: {e} (日付: {date_str})")
        return None

def move_and_rename(src_path, dest_dir, new_basename):
    """
    src_pathをdest_dir内にnew_basename + 元の拡張子で移動する。
    同名ファイルがある場合は、内容を比較して
        - 同一なら移動元のファイルを削除し、"duplicate" を返す。
        - 異なる場合は、上書きせず連番を付加して移動する。
    正常に移動できた場合は、"moved"、エラーが発生した場合は、"failed" を返す。
    """
    if not os.path.exists(src_path):
        print(f"移動元ファイルが見つかりません: {src_path}")
        return "failed" # 移動元がない
    ext = os.path.splitext(src_path)[1].lower()
    new_name = new_basename + ext
    dest_path = os.path.join(dest_dir, new_name)
    counter = 1
    # 同名ファイルが存在する場合の処理
    while os.path.exists(dest_path):
        # 内容を比較(shallow=Falseでバイナリデータを用いた内容の厳密な比較になる)
        try:
            if filecmp.cmp(src_path, dest_path, shallow=False):
                # 同一の内容の場合: 移動元ファイルを削除し、重複カウントを増やす
                print(f"重複削除: {src_path} (重複先: {dest_path})")
                print(f"重複ファイル検出: {src_path} と {dest_path} は同一の内容です。移動せずに {src_path} を削除します。")
                try:
                    os.remove(src_path)
                    return "duplicate" # ここで処理終了
                except OSError as e:
                    print(f"重複ファイルの削除失敗: {src_path} ({e})")
                    return "failed"
            else:
                # 異なる内容の場合: 連番を付与して新しい名前を作成
                new_name = f"{new_basename}_{counter}{ext}"
                dest_path = os.path.join(dest_dir, new_name)
                counter += 1
        except OSError as e: # ファイルアクセスエラーなど
            print(f"重複チェック/ファイルアクセスエラー: {e} (src: {src_path}, dest: {dest_path})")
            return "failed"
        except Exception as e:
            print(f"重複チェック中の予期せぬエラー: {e} (src: {src_path}, dest: {dest_path})")
            return "failed"
    try:
        shutil.move(src_path, dest_path)
        print(f"移動: {src_path} -> {dest_path}")
        return "moved"
    except Exception as e:
        print(f"移動失敗: {src_path} -> {dest_path} ({e})")
        return "failed"

def validate_and_parse_datetime(date_str):
    """
    撮影日時の文字列を受け取り、正しい日付としてパースします。
    タイムゾーン情報が含まれている場合は、東京(Asia/Tokyo)に合わせた後、tz情報を除去して返します。
    形式は "YYYY:MM:DD HH:MM:SS" または "YYYY:MM:DD HH:MM:SS+09:00" のような形式を想定します。
    """
    # 空白やNoneの場合
    if not date_str or not isinstance(date_str, str) or date_str.strip() == "" or "0000:00:00" in date_str:
        print(f"日付パース失敗: 入力が文字列ではありません (型: {type(date_str)})。")
        return None
    # 前処理: 不要な文字を除去
    date_str = re.sub(r'[^0-9:/\-T Z+.]', ' ', date_str).strip()
    parsed_dt = None
    original_tzinfo = None
    tz_match_found = None
    # 1. fromisoformatを試す
    try:
        iso_str = date_str.strip().replace(" ", "T")
        # マイクロ秒以降を切り捨て(最大6桁まで対応)
        if '.' in iso_str:
            parts = iso_str.split('.')
            iso_str = parts[0] + '.' + parts[1][:6]
        # Zを+00:00に
        if iso_str.endswith('Z'):
            iso_str = iso_str[:-1] + '+00:00'
        # タイムゾーンオフセットがない場合がある
        if 'T' in iso_str and not re.search(r'[+\-Z]', iso_str):
            # オフセットなしISO形式は naive としてパースされる
            pass
        # タイムゾーンオフセットの区切りがない場合(例: +0900) : を挿入
        iso_str = re.sub(r'([+\-])(\d{2})(\d{2})$', r'\1\2:\3', iso_str)
        temp_dt = datetime.fromisoformat(iso_str)
        parsed_dt = temp_dt
        original_tzinfo = parsed_dt.tzinfo # タイムゾーン情報を保持
    except (ValueError, TypeError):
        # 失敗したら次の方法へ
        pass
    # 2. strptimeで一般的なフォーマットを試す
    if not parsed_dt:
        formats_to_try = [
            "%Y:%m:%d %H:%M:%S",        # EXIF 標準
            "%Y-%m-%d %H:%M:%S",
            "%Y/%m/%d %H:%M:%S",
            # "%Y-%m-%dT%H:%M:%S",    # fromisoformatでカバーされるはず
            "%Y%m%d %H%M%S",          # 区切り文字なし
            "%Y:%m:%d %H:%M:%S.%f",    # マイクロ秒付き
            "%Y-%m-%d %H:%M:%S.%f",
            # タイムゾーン情報を含む可能性のある形式 (%zは限定的なので注意)
            # "%Y:%m:%d %H:%M:%S%z",
            # "%Y-%m-%d %H:%M:%S%z",
        ]
        # タイムゾーンらしき部分を除去してパースを試みる
        cleaned_str = date_str.strip()
        tz_match = re.search(r'([+\-]\d{2}:?\d{2}|Z)\s*$', cleaned_str)
        if tz_match:
            tz_match_found = tz_match.group(1)
            cleaned_str = cleaned_str[:tz_match.start()].strip()
        for fmt in formats_to_try:
            try:
                # マイクロ秒を含むフォーマットの場合、入力にマイクロ秒がなければエラーになるため調整 
                fmt_to_use = fmt
                input_to_use = cleaned_str
                if ".%f" in fmt and '.' not in cleaned_str:
                    # フォーマットから.%fを除去
                    fmt_to_use = fmt.replace(".%f", "")
                elif ".%f" not in fmt and '.' in cleaned_str:
                    # 入力から.%f以降を除去
                    input_to_use = cleaned_str.split('.')[0]
                temp_dt = datetime.strptime(input_to_use, fmt_to_use)
                parsed_dt = temp_dt
                break # パース成功
            except ValueError:
                continue
    if not parsed_dt:
        print(f"日付パース失敗: 入力 '{date_str}' は既知のフォーマットに一致しませんでした。")
        return False
    # 3. タイムゾーン処理(Asia/Tokyo基準のnaive datetimeにする)
    try:
        tokyo_tz = pytz.timezone('Asia/Tokyo')
        dt_final_naive = None
        if original_tzinfo: # fromisoformatでタイムゾーンが取得できた場合
            dt_aware = parsed_dt # すでにaware
            dt_tokyo = dt_aware.astimezone(tokyo_tz)
            dt_final_naive = dt_tokyo.replace(tzinfo=None)
        elif tz_match_found: # strptimeでパースし、タイムゾーンらしき文字列が見つかった場合
            tz_str = tz_match_found
            fixed_tz = None
            try:
                if tz_str == 'Z':
                    fixed_tz = pytz.utc
                elif ':' in tz_str:
                    fixed_tz = pytz.FixedOffset(int(tz_str[-5:-3])*60 + int(tz_str[-2:]) * (1 if tz_str.startswith('+') else -1))
                elif len(tz_str) == 5: # +HHMM
                    fixed_tz = pytz.FixedOffset(int(tz_str[1:3])*60 + int(tz_str[3:5]) * (1 if tz_str.startswith('+') else -1))
                if fixed_tz:
                    aware_dt = fixed_tz.localize(parsed_dt) # naiveをawareに
                    dt_tokyo = aware_dt.astimezone(tokyo_tz)
                    dt_final_naive = dt_tokyo.replace(tzinfo=None)
                else: # オフセットが不明な場合はローカルタイムと仮定
                    aware_dt = tokyo_tz.localize(parsed_dt, is_dst=None)
                    dt_final_naive = aware_dt.replace(tzinfo=None) # 既に東京時間なのでtzinfo除去のみ
            except Exception as tz_err:
                print(f"警告: タイムゾーン '{tz_str}' の処理中にエラー({tz_err})。ローカルタイムと仮定します。")
                aware_dt = tokyo_tz.localize(parsed_dt, is_dst=None)
                dt_final_naive = aware_dt.replace(tzinfo=None)
        else:
            # Naive datetime は元々がAsia/Tokyoであると仮定
            try:
                aware_dt = tokyo_tz.localize(parsed_dt, is_dst=None)
                dt_final_naive = aware_dt.replace(tzinfo=None) # tzinfo除去
            except (pytz.AmbiguousTimeError, pytz.NonExistentTimeError) as e:
                print(f"警告: ローカル時刻 '{parsed_dt}' のタイムゾーン割り当て時に問題 ({e})。そのまま naive な時刻を使用します。")
                dt_final_naive = parsed_dt # naiveのまま使用
        # 4. 有効範囲チェック
        min_allowed_date = datetime(1970, 1, 1)
        max_allowed_date = datetime.now() + timedelta(days=365) # 1年先まで許容
        if not (min_allowed_date <= dt_final_naive <= max_allowed_date):
            print(f"日付範囲外エラー: 処理後の日付 '{dt_final_naive}' は許容範囲外です (元: '{date_str}')。")
            return False
        return dt_final_naive # 成功時は naive な datetime オブジェクトを返す
    except pytz.UnknownTimeZoneError:
        print("タイムゾーン 'Asia/Tokyo' がみつかりません。pytzを確認してください。")
        return False
    except Exception as e:
        print(f"タイムゾーン処理/日付処理中に予期せぬエラー: {e} (入力: '{date_str}', パース結果: {parsed_dt})")
        return False
    

def count_media_files(source_folder):
    """メディアファイルの数をカウント"""
    image_count = 0
    video_count = 0
    metadata_count = 0
    other_count = 0
    for dirpath, _, filenames in os.walk(source_folder):
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext in IMAGE_EXTS:
                image_count += 1
            elif ext in VIDEO_EXTS:
                video_count += 1
            elif ext in METADATA_EXTS:
                metadata_count += 1
            else:
                other_count += 1 # 対象外ファイルもカウント
    return image_count, video_count, metadata_count, other_count

def thread_count():
    """環境に応じた最適なスレッド数を返す"""
    cpu_count = multiprocessing.cpu_count()
    if platform.system() == 'Darwin':
        # is_apple_silicon = platform.processor() == 'arm' or platform.machine().startswith('arm') # より確実な判定
        # Apple Silicon の判定を修正 (processor()は空を返すことがあるためmachine()を見る)
        is_apple_silicon = 'arm' in platform.machine().lower()
        if is_apple_silicon:
            # 性能コア優先で少し多めに (最大16程度)
            optimal_threads = max(4, min(int(cpu_count * 0.5), 16))
            print(f"Apple Silicon検出: スレッド数 = {optimal_threads}")
            return optimal_threads
        else: # Intel Mac
            optimal_threads = max(4, min(cpu_count * 2, 16))
            print(f"Intel Mac検出: スレッド数 = {optimal_threads}")
            return optimal_threads
    else: # Windows, Linuxなど
        optimal_threads = max(4, min(cpu_count * 2, 16))
        print(f"{platform.system()}検出: スレッド数 = {optimal_threads}")
        return optimal_threads

# --- 非同期処理(async/await) ---
async def async_process_file(file_path, dest_root, loop, executor):
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in IMAGE_EXTS and ext not in VIDEO_EXTS:
        return {"moved": 0, "duplicate": 0, "failed": 0}
    # ブロッキングなget_file_dateもrun_in_executorで呼ぶ
    date_str = await loop.run_in_executor(executor, get_file_date, file_path)
    if not date_str:
        print(f"日付情報なし: {file_path} をスキップします。")
        return {"moved": 0, "duplicate": 0, "failed": 0}
    dest_info = await loop.run_in_executor(executor, make_destination_path, dest_root, date_str)
    if not dest_info:
        print(f"移動先ディレクトリ作成失敗: {file_path} をスキップします。")
        return {"moved": 0, "duplicate": 0, "failed": 0}
    dest_dir, new_basename = dest_info
    # asyncio.Lockで同じディレクトリへの同時アクセスを防ぐ
    async with dir_locks_lock:
        dir_lock = dir_locks.setdefault(dest_dir, asyncio.Lock())
    result_counts = {"moved": 0, "duplicate": 0, "failed": 0}
    async with dir_lock:
        # メインのメディアファイルを移動/リネーム
        res = await loop.run_in_executor(executor, move_and_rename, file_path, dest_dir, new_basename)
        if res in result_counts:
            result_counts[res] += 1
        else:
            result_counts["failed"] += 1 # 不明な場合は失敗
        # メインファイルが移動または重複削除された場合、関連ファイルを処理
        if res == "moved" or res == "duplicate":
            if ext in VIDEO_EXTS:
                original_basename_no_ext = os.path.splitext(os.path.basename(file_path))[0]
                original_dir = os.path.dirname(file_path)
                # .thm ファイル (元のファイル名.thm)
                thm_filename = original_basename_no_ext + ".thm"
                thm_path = os.path.join(original_dir, thm_filename)
                if os.path.exists(thm_path):
                    # .thm の移動結果はメインの結果に含めない（個別にカウントしない）
                    thm_res = await loop.run_in_executor(executor, move_and_rename, thm_path, dest_dir, new_basename)
                    if thm_res == "failed":
                        print(f"警告: 関連THM {thm_filename} の移動失敗")
                # .xml ファイル (命名規則を複数試す)
                # core + suffix + M01.xml, core + M01 + suffix + .xml, basename.xml など 
                m = re.match(r'^(.*?)(\s*\(_?\d+\))?$', original_basename_no_ext) # _1 や (1) に対応
                core = original_basename_no_ext
                suffix = ''
                if m:
                    core = m.group(1)
                    suffix = m.group(2) or ''
                xml_patterns = [
                    f"{core}{suffix}M01.xml", # MyVideo (1)M01.xml
                    f"{core}M01{suffix}.xml", # C0029M01 (1).xml
                    f"{original_basename_no_ext}.xml", # MyVideo (1).xml
                    f"{original_basename_no_ext}.XML", # 大文字
                    f"{core}{suffix}M01.XML",
                    f"{core}M01{suffix}.XML",
                ]
                found_xml_path = None
                xml_filename_found = None
                for pattern in xml_patterns:
                    potential_xml_path = os.path.join(original_dir, pattern)
                    if os.path.exists(potential_xml_path):
                        found_xml_path = potential_xml_path
                        xml_filename_found = pattern
                        break
                if found_xml_path:
                    # .xml の移動結果もメインの結果に含めない
                    xml_res = await loop.run_in_executor(executor, move_and_rename, found_xml_path, dest_dir, new_basename)
                    if xml_res == "failed":
                        print(f"警告: 関連XML {xml_filename_found} の移動失敗")

    return result_counts

async def async_main(source_folder, dest_root):
    loop = asyncio.get_running_loop()
    num_threads = thread_count()
    executor = ThreadPoolExecutor(max_workers=num_threads)
    # 対象ファイルのパスをリストアップ(これは軽い処理なので同期でOK)
    file_list = []
    print("ファイルリスト作成中...")
    for dirpath, _, filenames in os.walk(source_folder):
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext in IMAGE_EXTS or ext in VIDEO_EXTS:
                file_list.append(os.path.join(dirpath, filename))
    print(f"{len(file_list)} 個のメディアファイルを検出しました。")
    tasks = [async_process_file(file_path, dest_root, loop, executor) for file_path in file_list]
    # 進捗表示
    results = []
    processed_count = 0
    total_files = len(tasks)
    print("ファイル処理を開始します...")
    for future in asyncio.as_completed(tasks):
        result = await future
        results.append(result)
        processed_count += 1
        # コンソールに進捗を表示
        if processed_count % 10 == 0 or processed_count == total_files:
            print(f"進捗: {processed_count}/{total_files} ファイル処理完了")
    total_moved = sum(r["moved"] for r in results)
    total_duplicate = sum(r["duplicate"] for r in results)
    total_failed = sum(r["failed"] for r in results)
    # スキップされたファイル数 (日付なし or 移動先作成失敗)
    total_skipped = total_files - (total_moved + total_duplicate + total_failed)
    print("全ファイルの処理が完了しました。")
    executor.shutdown(wait=True) # Executorをシャットダウン
    return total_moved, total_duplicate, total_failed, total_skipped, total_files

def gui_main():
    # ロケール設定 (エラーハンドリング付き)
    try:
        # Windowsの場合は 'ja-JP' や 'japanese' も試す
        locales_to_try = ['ja_JP.UTF-8', 'ja_JP.utf8', 'Japanese_Japan.932', 'ja-JP', 'japanese', '']
        set_locale = False
        for loc in locales_to_try:
            try:
                locale.setlocale(locale.LC_ALL, loc)
                print(f"ロケールを '{loc}' に設定しました。")
                set_locale = True
                break
            except locale.Error:
                continue
        if not set_locale:
            print("警告: 日本語ロケールの設定に失敗しました。デフォルトロケールを使用します。")
    except Exception as e:
         print(f"ロケール設定中に予期せぬエラー: {e}")
    # Tkinterのルートウィンドウを作成（表示はしない）
    root = tk.Tk()
    root.withdraw()

    # --- ソースフォルダー選択 ---
    source_folder = filedialog.askdirectory(title="整理対象のフォルダーを選択")
    if not source_folder:
        messagebox.showerror("エラー", "ソースフォルダーが選択されませんでした。")
        return

    # --- 移動先フォルダー選択 ---
    dest_root = filedialog.askdirectory(title="移動先のフォルダーを選択")
    if not dest_root:
        messagebox.showerror("エラー", "移動先フォルダーが選択されませんでした。")
        return
    
    print("処理対象のファイル数をカウントしています...")
    # メディアファイル数のカウント
    image_count, video_count, metadata_count, other_count = count_media_files(source_folder)
    total_media_files = image_count + video_count

    count_message = (
        f"【処理対象ファイル】\n"
        f"  画像: {image_count} 個\n"
        f"  動画: {video_count} 個\n"
        f"  合計: {total_media_files} 個\n"
        f"--------------------\n"
        f"【関連ファイル（同時に移動されます）】\n"
        f"  メタデータ (.xml, .thm): {metadata_count} 個\n"
        f"--------------------\n"
        f"【対象外ファイル（移動されません）】\n"
        f"  その他: {other_count} 個"
    )
    print(count_message) # コンソールにも表示

    if not messagebox.askyesno("処理内容の確認", f"{count_message}\n\nこれらのメディアファイル ({total_media_files}個) を撮影日時に基づいて\n「{dest_root}」\nに整理しますか？"):
        print("処理はキャンセルされました。")
        return
    # --- 非同期処理の実行 ---
    print("非同期処理を実行します...")
    # asyncio.run() は Windows で SelectorEventLoop を使う
    # ProactorEventLoop が必要な場合がある (特に subprocess 関連)
    if platform.system() == "Windows":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    total_moved, total_duplicate, total_failed, total_skipped, total_processed = asyncio.run(async_main(source_folder, dest_root))

    # --- 結果表示 ---
    summary = (
        f"【処理結果】\n"
        f"  処理対象ファイル数: {total_processed} 件\n"
        f"--------------------\n"
        f"  移動成功: {total_moved} 件\n"
        f"  重複ファイル(削除): {total_duplicate} 件\n"
        f"  スキップ(日付なし等): {total_skipped} 件\n"
        f"  移動失敗: {total_failed} 件\n"
        f"--------------------\n"
        f"  (成功 + 重複 + スキップ + 失敗 = {total_moved + total_duplicate + total_skipped + total_failed})"
    )
    print(summary) # コンソールにも表示
    messagebox.showinfo("処理完了", summary)

if __name__ == "__main__":
    # 必要に応じてExifToolのパスを指定
    # exiftool_path = r"C:\path\to\exiftool.exe" # Windowsの例
    # exiftool_path = "/usr/local/bin/exiftool" # macOS/Linuxの例
    # if 'exiftool_path' in locals() and os.path.exists(exiftool_path):
    #     exiftool.ExifToolHelper.executable = exiftool_path
    #     print(f"ExifToolのパスを '{exiftool_path}' に設定しました。")
    # else:
    #     print("システムPATHにあるExifToolを使用します。")
    #     # パスが通っているか簡単なチェック
    #     try:
    #         with exiftool.ExifToolHelper() as et:
    #             print(f"ExifToolバージョン: {et.version}")
    #     except exiftool.ExifToolExecuteError:
    #         print("警告: ExifToolが見つからないか、実行できません。パスを確認してください。")
    #     except Exception as e:
    #         print(f"ExifToolのチェック中にエラー: {e}")
    gui_main()
