import os
import shutil
import time
import datetime
from datetime import date,datetime
import numpy as np
import pandas as pd

FILE_DICT = {'Шаблон':['',
                       'J:/~ 09_Машинное обучение_Прогноз показателей СЭР/ЧЕК-ЛИСТЫ и DATA-SHOP/DATA-SHOP/!DATA-SHOP! — шаблон.xlsx',
                       'Для витрин'],
             'Евсина': ['Отдел развития социальной сферы',
                        'J:/~ 05_Отдел развития социальной сферы/ПРОЕКТЫ/!DATA-SHOP!.xlsx',
                        'Лист1'],
             'Оганесян': ['Отдел продвинутой аналитики и машинного обучения',
                          'J:/~ 09_Машинное обучение_Прогноз показателей СЭР/ЧЕК-ЛИСТЫ и DATA-SHOP/DATA-SHOP/!DATA-SHOP! - PYTHONUS.xlsx',
                          'Для витрин'],
             'Окунькова': ['Отдел комлексного мониторинга',
                           'J:/~ 08_Отдел комлексного мониторинга/Политика обращения с данными/!DATA-SHOP!.xlsx',
                           'Лист1'],
             'Пушилин': ['Отдел развития реального сектора экономики',
                         'J:/~ 04_Отдел развития реального сектора экономики/~ Документы отдела/!DATA-SHOP!.xlsm',
                         'Лист1'],
             'Шибина': ['Отдел аграрной и продовольственной политики',
                        'J:/~ 03_Отдел аграрной и продовольственной политики/16 НОВЫЕ ПРОЕКТЫ/!DATA-SHOP! - отдел АПК.xlsx',
                        'Для витрин']}

for keys in list(FILE_DICT.keys()):
    if keys == 'Шаблон':
        Department = FILE_DICT[keys][0]
        file_name = FILE_DICT[keys][1]
        sheet_name = FILE_DICT[keys][2]
        RESULTS = pd.read_excel(file_name, sheet_name=sheet_name, header=0)
        RESULTS['department'] = ''
        RESULTS.loc[2, 'department'] = 'Отдел'
    else:
        Department = FILE_DICT[keys][0]
        file_name = FILE_DICT[keys][1]
        sheet_name = FILE_DICT[keys][2]
        temp = pd.read_excel(file_name, sheet_name=sheet_name, header=3)
        temp['department'] = Department
        RESULTS = pd.concat([RESULTS, temp])

RESULTS.index = range(RESULTS.shape[0])
RESULTS

File_name = r'J:\~ 09_Машинное обучение_Прогноз показателей СЭР\ЧЕК-ЛИСТЫ и DATA-SHOP\DATA-SHOP/DATA-SHOP_сводный по отделам.xlsx'
Sheet_name = str(date.today())

with pd.ExcelWriter(File_name, engine="openpyxl", mode='a') as writer:
    RESULTS.to_excel(writer, sheet_name=Sheet_name, header=True, index=False, encoding='1251')


