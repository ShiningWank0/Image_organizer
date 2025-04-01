import exiftool
import re

# files = ["/Users/taku/Desktop/USBなどからデータコピー/Pixel7/Camera/PXL_20250316_011505091.mp4", "/Users/taku/Desktop/USBなどからデータコピー/カメラ/C0001.MP4", "/Users/taku/Desktop/USBなどからデータコピー/102MSDCF/MOV00355.MPG"]
files = ["/Users/taku/Desktop/USBなどからデータコピー/Camera/VID_20200531_104522.mp4"]
# files = ["/Users/taku/Desktop/test/GEDC0040.AVI", "/Users/taku/Desktop/test/GEDC0049.AVI"]
# files = ["/Users/taku/Desktop/USBなどからデータコピー/カメラ1/GEDC0077.AVI", "/Users/taku/Desktop/USBなどからデータコピー/カメラ1/GEDC0050.AVI", "/Users/taku/Desktop/USBなどからデータコピー/カメラ1/GEDC0078.AVI", "/Users/taku/Desktop/USBなどからデータコピー/カメラ1/GEDC2002.AVI"]
# files = ["/Users/taku/Desktop/USBなどからデータコピー/101MSDCF/DSC01357.JPG"]
# files = ["/Users/taku/Desktop/USBなどからデータコピー/カメラ1/GEDC0040.AVI"]
# files = ["/Users/taku/Desktop/USBなどからデータコピー/カメラ4/C0019.MP4"]

with exiftool.ExifToolHelper() as et:
    metadata = et.get_metadata(files)
    for d in metadata:
        for k, v in d.items():
            print(f"{k} = {v}")
            # if "DateTimeOriginal" in k:
            #     result = re.sub(r'([+\-]\d{2}:\d{2})$', '', v)
            #     print(result)