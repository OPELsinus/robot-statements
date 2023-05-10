import datetime
from pathlib import Path

from tools import hold_session
import time


# for i in range(1000):
#     hold_session()
#     time.sleep(400)
import os
import shutil
# shutil.rmtree(rf'\\vault.magnum.local\Common\Stuff\_05_Финансовый Департамент\01. Казначейство\Сверка\Сверка РОБОТ\Temp1111.xlsx')
# os.remove(rf'\\vault.magnum.local\Common\Stuff\_05_Финансовый Департамент\01. Казначейство\Сверка\Сверка РОБОТ\Temp1111.xlsx')
print(datetime.date.today().strftime('%d.%m.%Y')[-4:])
#
# book = openpyxl.load_workbook(r'C:\Users\Abdykarim.D\Desktop\аф29.xlsx')
# sheet = book.active
#
# rows_to_delete = []
#
# for row in sheet.iter_rows(min_row=9):
#     if row[2].value is None:
#         break
#
#     if 'эконом' not in row[2].value.lower():
#         rows_to_delete.append(row[0].row)
#
# for row_num in reversed(rows_to_delete):
#     sheet.delete_rows(row_num)
#
# for row in sheet.iter_rows(min_row=9):
#     if row[3].value is None:
#         break
#
#     if 'подсолн' in row[3].value.lower() and ' 5л' in row[3].value.lower():
#         print(row[17].value)
#         row[17].value *= 5
#
# book.save(r'C:\Users\Abdykarim.D\Desktop\аф29999.xlsx')
