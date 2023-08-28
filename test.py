import os
import shutil

for file in os.listdir(r'\\vault.magnum.local\Common\Stuff\_05_Финансовый Департамент\01. Казначейство\Сверка\Сверка РОБОТ\июнь 2023'):
    # print(file)
    if '_1' in file:
        file_ = file.split('_1')[0] + file.split('_1')[1]
        print(file_)
        shutil.move(fr'\\vault.magnum.local\Common\Stuff\_05_Финансовый Департамент\01. Казначейство\Сверка\Сверка РОБОТ\июнь 2023\{file}', fr'\\vault.magnum.local\Common\Stuff\_05_Финансовый Департамент\01. Казначейство\Сверка\Сверка РОБОТ\июнь 2023\{file_}')
