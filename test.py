import os
import shutil
import traceback
from pathlib import Path
from zipfile import ZipFile
import os

from config import halyk_extract_path

day = '09.01.2024'

for folder, subfolder, files in os.walk(halyk_extract_path):
    print(files)
    print(folder)
    print(subfolder)
    print('-')
    for file in files:
        print(file)
        if day in file and 'kzt народный' in file.lower() and os.path.getsize(os.path.join(folder, file)) / 1024 > 100:
            break

    print('-----')
