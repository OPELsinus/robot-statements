import os
import shutil
import traceback
from pathlib import Path
from zipfile import ZipFile
import os

try:
    file_path = Path(r'C:\Users\Abdykarim.D\Documents\График_инвентаризаций_y2023.xlsx')
    tmp_folder = file_path.parent.joinpath('__temp__')
    with open(file_path, 'rb') as excel_file:
        with ZipFile(excel_file) as excel_container:
            excel_container.extractall(tmp_folder)
            excel_container.close()
    wrong_file_path = os.path.join(tmp_folder.__str__(), 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder.__str__(), 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)
    file_path.unlink()
    shutil.make_archive(file_path.__str__(), 'zip', tmp_folder)
    os.rename(file_path.__str__() + '.zip', file_path.__str__())
    shutil.rmtree(tmp_folder.__str__(), ignore_errors=True)
except (Exception,):
    traceback.print_exc()
