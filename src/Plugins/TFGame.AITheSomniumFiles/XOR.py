import subprocess
import os
from pathlib import Path

def xor(input_file, output_file):
    try:
        with open(input_file, 'rb') as f:
            data = bytearray(f.read())

        for i in range(4, len(data)):
            data[i] ^= i % 256

        with open(output_file, 'wb') as f:
            f.write(data)
    except Exception as e:
        print(f"Error processing file {input_file}: {e}")

def process_files_in_directory(directory):
    for path in Path(directory).rglob('*.bytes'):
        if path.is_file():
            input_file = str(path)
            output_file = os.path.join(os.path.dirname(input_file), 'mynewfile.lua')
            xor(input_file, output_file)

def main():
    directory = path("F:/Xbox/AI- THE SOMNIUM FILES/Content/AI_TheSomniumFiles_Data/StreamingAssets/AssetBundles/StandaloneWindows64/TextAsset/")
    # Change this to the desired directory
    process_files_in_directory(directory)
    
def path(arg):
    return arg.replace("\\", "/")

if __name__ == "__main__":
    main()
