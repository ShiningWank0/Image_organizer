import tkinter as tk
from tkinter import filedialog
import ffmpeg
import os
from datetime import datetime, timedelta
import pytz # ffmpegで抽出するとタイムゾーンの違いにより正常な時間が表示されないことの対策

def load_video_metadata(local_timezone='Asia/Tokyo'):
    # Create and hide the Tkinter root windwo
    root = tk.Tk()
    root.withdraw()

    # Open file dialog to select video file
    file_path = filedialog.askopenfilename(
        title="Select Video File",
        filetypes=[
            ("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm")
        ]
    )

    if file_path:
        try:
            print(file_path)
            # Use ffprobe to extract metadata
            probe = ffmpeg.probe(file_path)
            # Extract metadata
            metadata = {}
            # General video stream information
            video_info = next((stream for stream in probe['streams'] if stream['codec_type'] == 'video'), None)
            if video_info:
                metadata['Video'] = {
                    'Codec': video_info.get('codec_name', 'Unknown'),
                    'Width': video_info.get('width', 'Unknown'),
                    'Height': video_info.get('height', 'Unknown'),
                    'Duration': video_info.get('duration', 'Unknown')
                }
            
            # Audio stream information
            audio_info = next((stream for stream in probe['streams'] if stream['codec_type'] == 'audio'), None)
            if audio_info:
                metadata['Audio'] = {
                    'Codec': audio_info.get('codec_name', 'Unknown'),
                    'Channels': audio_info.get('channels', 'Unknown'),
                    'Sample Rate': audio_info.get('sample_rate', 'Unknown')
                }
            
            # Format information (including creation time)
            if 'format' in probe:
                format_info = probe['format']
                metadata['Format'] = {
                    'Format Name': format_info.get('format_name', 'Unknown'),
                    'Size': format_info.get('size', 'Unknown') + ' bytes',
                    'Bit Rate': format_info.get('bit_rate', 'Unknown')
                }
                # Try to extract creation time
                creation_time = format_info.get('tags', {}).get('creation_time')
                if creation_time:
                    try:
                        # UTCタイムスタンプをパース
                        utc_time = datetime.fromisoformat(creation_time.replace('Z', '+00:00'))
                        # 指定されたローカルタイムゾーンに変換
                        local_tz = pytz.timezone(local_timezone)
                        local_time = utc_time.astimezone(local_tz)
                        metadata['Timestamp'] = {
                            'Creation Time (UTC)': utc_time.strftime('%Y-%m-%d %H:%M:%S'),
                            f'Creation Time({local_timezone})': local_time.strftime('%Y-%m-%d %H:%M:%S'),
                            'Formatted': format_datetime(local_time.strftime('%Y:%m:%d %H:%M:%S'))
                        }
                    except Exception as time_error:
                        metadata['Timestamp'] = {
                            'Raw Creation Time': creation_time,
                            'Error': str(time_error)
                        }
            
            return metadata
        
        except Exception as e:
            print(f"An error occurred: {e}")
            return None
    else:
        print("No file was selected.")
        return None

def format_datetime(datetime_str):
    # Convert datetime string from 'YYYY:MM:DD HH:MM:SS' to 'YYYY_MM_DD_HH_MM_SS'
    try:
        return datetime_str.replace(':', '_').replace(' ', '_')
    except:
        return datetime_str

def pretty_print_metadata(metadata):
    if metadata is None:
        return
    
    for category, details in metadata.items():
        print(f"\n--- {category} ---")
        for key, value in details.items():
            print(f"{key}: {value}")

def main():
    # Main processing
    video_metadata = load_video_metadata()
    pretty_print_metadata(video_metadata)

if __name__ == "__main__":
    main()