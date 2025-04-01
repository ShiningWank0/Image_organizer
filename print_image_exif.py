import tkinter as tk
from tkinter import filedialog
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
import os
import base64

def load_exif():
    # Tkinterのルートウィンドウを作成し、非表示にする
    root = tk.Tk()
    root.withdraw()

    # デフォルトのファイルパスを設定（ユーザーのピクチャディレクトリ）
    # default_path = os.path.expanduser("~/Pictures")

    # ファイル選択ダイアログを表示
    file_path = filedialog.askopenfilename(
        title="画像を選択",
        filetypes=[("Image files", "*.jpg *.jpeg *.png *.tif *.tiff *.heic")],
        # initialdir=default_path
    )

    if file_path:
        try:
            print(file_path)
            ifd_dict = {}
            with Image.open(file_path) as im:
                exif = im.getexif()
            
            # Zeroth IFD (Image File Directory)
            ifd_dict["Zeroth"] = {}
            for tag_id, value in exif.items():
                tag_name = TAGS.get(tag_id, hex(tag_id))
                decoded_value = decode_value(value)
                if decoded_value is not None:
                    ifd_dict["Zeroth"][tag_name] = decoded_value

            # Exif IFD
            if 0x8769 in exif:
                ifd_dict["Exif"] = {}
                for tag_id, value in exif.get_ifd(0x8769).items():
                    tag_name = TAGS.get(tag_id, hex(tag_id))
                    decoded_value = decode_value(value)
                    if decoded_value is not None:
                        ifd_dict["Exif"][tag_name] = decoded_value

            # GPS IFD
            if 0x8825 in exif:
                ifd_dict["GPSInfo"] = {}
                for tag_id, value in exif.get_ifd(0x8825).items():
                    tag_name = GPSTAGS.get(tag_id, hex(tag_id))
                    decoded_value = decode_value(value)
                    if decoded_value is not None:
                        ifd_dict["GPSInfo"][tag_name] = decoded_value

            return ifd_dict
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return None
    else:
        print("ファイルが選択されませんでした。")
        return None

def decode_value(value):
    # NoneやEmptryの値を処理
    if value is None:
        return None
    
    # バイト列の処理
    if isinstance(value, bytes):
        # 印字可能な文字列かどうかをチェック
        try:
            # UTF-8, UTF-16, ASCII, Latin-1で順次試行
            decode_attempts = [
                ("utf-8", "strict"),
                ("utf-16le", "replace"),
                ("utf-16be", "replace"),
                ("ascii", "replace"),
                ("latin-1", "replace"),
                ("mac-roman", "replace")
            ]
            for encoding, errors in decode_attempts:
                try:
                    decoded = value.decode(encoding, errors=errors)
                    # 印字可能な文字のみを含む場合に返す
                    if all(32 <= ord(c) < 127 or c.isprintable() for c in decoded):
                        return decoded.strip()
                except:
                    continue
            # デコードできない場合は、Base64エンコードして返す
            return base64.b64encode(value).decode("ascii")
        except Exception as e:
            # 最終的にデコードできない場合
            return f"[Undecoded Data: {len(value)} bytes]"
    # 他の型の値をそのまま返す
    return value

def format_datetime(datetime_str):
    # DateTimeOriginalの文字列を指定のフォーマットに変換
    try:
        # フォーマット: 2019:08:26 09:54:50 → 2019_08_26_09_54_50
        return datetime_str.replace(':', '_').replace(' ', '_')
    except:
        return datetime_str

def pretty_print_exif(exif_dict):
    if exif_dict is None:
        return
    
    for ifd_type, tags in exif_dict.items():
        print(f"\n--- {ifd_type} IFD ---")
        for tag_name, tag_value in sorted(tags.items()):
            # 長すぎる値は省略して表示
            if isinstance(tag_value, str) and len(tag_value) > 200:
                print(f"{tag_name}: {tag_value[:200]}... (全{len(tag_value)}文字)")
            else:
                print(f"{tag_name}: {tag_value}")

def main():
    # メイン処理
    ifd_dict = load_exif()
    pretty_print_exif(ifd_dict)
    if ifd_dict and 'Exif' in ifd_dict:
        # DateTimeOriginalを取得して指定のフォーマットで表示
        datetime_original = ifd_dict['Exif'].get('DateTimeOriginal')
        if datetime_original:
            formatted_datetime = format_datetime(datetime_original)
            print(formatted_datetime)
        else:
            print("DateTimeOriginalが見つかりませんでした。")

if __name__ == "__main__":
    main()