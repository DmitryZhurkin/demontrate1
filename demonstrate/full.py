import datetime
from pandastable import Table
import tkinter as tk
from PIL import Image, ImageDraw, ImageFont
import PySimpleGUI as sg
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import numpy
import re
import itertools
import shutil
import openpyxl
from xml.etree.ElementTree import Element
from tqdm import tqdm
import os
import warnings

import warnings

warnings.filterwarnings("ignore")
# import xml.etree.ElementTree
working_directory = os.getcwd()

sg.theme("SystemDefaultForReal")
months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
          "Декабрь"]
insertions = ['январе', 'феврале', 'марте', 'апреле', 'мае', 'июне', 'июле', 'августе', 'сентябре', 'октябре', 'ноябре',
              'декабре']
c = []
dirname = os.path.dirname(__file__)
dirname1 = os.path.dirname(__file__) + "/sluc/"
xlsx = pd.read_excel(dirname1 + "sluckii22.xlsx")
layout = [  # [sg.Text('Window normal', size=(30, 1), key='Status')],
    # [sg.Text("Выберите файл:")],
    # [sg.InputText(size=(24,1),key="-FILE_PATH-"),
    #  sg.FileBrowse('Выбрать ',
    #                 initial_folder="Z:/данные из ГИС ОМС/УД",
    #                file_types=[("xlsx Files", "*.xlsx"),("xlsm Files", "*.xlsm") ])],
     [sg.Text('Углубленная диспанцеризация', font=1, justification='Center', expand_x=True, expand_y=True)],
    [sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]), sg.InputText(size=(25, 1), key="-FILE_PATH_DISP-", expand_x=True, expand_y=True)],
    [sg.Button("1. Первичная обработка", size=(27, 2), key = '-disp-', font=20,  expand_x=True, expand_y=True)],
    [sg.Button("2. Сформировать письма по 29 признаку", size=(30, 2), key = 'let29', font=20, expand_x=True, expand_y=True)],
    [sg.Button("2. Сформировать письма по 31 признаку", size=(30, 2), key='let31', font=20, expand_x=True, expand_y=True)],
[sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС/УД", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]), sg.InputText(size=(25, 1), key="-FILE_PATH_VIBER-", expand_x=True, expand_y=True)],
    [sg.Button("3. Загрузить viber-уведомления", size=(30, 2), font=30, key='Viber', expand_x=True, expand_y=True)],
    # sg.Submit("3. Загрузить viber-уведомления", size=(27, 2), font=20)],
    # [sg.Text("Прогресс обработки по письмам:            ")], #sg.ProgressBar(len(xlsx), orientation='h', size=(23, 23), key='progressbar')],
[sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС/УД", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]), sg.InputText(size=(25, 1), key="-FILE_PATH_SMS-", expand_x=True, expand_y=True)],
    [sg.Button("4. Загрузить смс-уведомления", size=(30, 2), key='Sms',font=30, expand_x=True, expand_y=True)],
             [sg.FileBrowse('Выбрать файл',
                           initial_folder="Z:/данные из ГИС ОМС/УД", font=1,
                           file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]) ,   sg.InputText(size=(25, 1), key="-FILE_PATH_BASE-",  expand_x=True, expand_y=True),],
    [sg.Button("5. Загрузить в базу", size=(30, 2), font=30,key='Base', expand_x=True, expand_y=True)],
    [sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС/УД", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]),
     sg.InputText(size=(25, 1), key="-FILE_PATH_INDEXDISP-", expand_x=True, expand_y=True)],
    [sg.Button("Поменять порядок записей", size=(30, 2), key = 'idisp', font=30, expand_x=True, expand_y=True)],
[sg.Text('Профосмотр', size=(50, 1), font=50,justification='Center',expand_x=True, expand_y=True)],
[sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]),
    sg.InputText(size=(25, 1), key="-FILE_PATH_PROF-", expand_x=True, expand_y=True)],
    [sg.Button("Сформировать письма по списку", size=(30, 2), key = 'prof',font=30, expand_x=True, expand_y=True, button_color= 'blue')],

[sg.Text('Вакцинация', size=(30, 1), font=30,justification='Center',expand_x=True, expand_y=True)],
[sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]),
    sg.InputText(size=(25, 1), key="-FILE_PATH_VACT-", expand_x=True, expand_y=True)],
    [sg.Button("Обработка по вакцинации", size=(30, 2), font=30, key = 'vact', expand_x=True, expand_y=True, button_color='red')],
    [sg.FileBrowse('Выбрать файл',
                   initial_folder="Z:/данные из ГИС ОМС/Вакцинация", font=1,
                   file_types=[("xlsx Files", "*.xlsx"), ("xlsm Files", "*.xlsm")]),
    sg.InputText(size=(25, 1), key="-FILE_PATH_INDEXVACT-", expand_x=True, expand_y=True)],
    [sg.Button("Поменять порядок записей", size=(30, 2), key = 'ivact', font=30, expand_x=True, expand_y=True, button_color='red')],
]

window = sg.Window('PolisInform', icon='1.ico', resizable=True, finalize=True).Layout(layout)
# window = sg.Window('PolisInform', icon='1.ico').Layout(layout)
sg.SetOptions(text_justification='left')
while True:  # The Event Loop
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'Configure':
        if window.TKroot.state() == 'zoomed':
            status.update(value='Window zoomed and maximized !')
        else:
            status.update(value='Window normal')
    if event == '-disp-':
        start_time = datetime.now()

        # server = "DMITRY\SQLEXPRESS"
        server = "10.58.1.200"
        dbname = "RGSNEW1"
        uname = "rgs"
        pword = "233239"
        # eng = create_engine("mssql+pyodbc://"+server+"/"+dbname+"?driver=SQL+Server")
        eng = create_engine(
            "mssql+pyodbc://" + uname + ":" + pword + "@" + server + "/" + dbname + "?driver=SQL+Server")
        pathh = values["-FILE_PATH_DISP-"]
        xlsx = pd.read_excel(pathh, dtype={"A6": str, "A8": str})
        xlsx = pd.read_excel(pathh, header=4, names=(
            "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15",
            "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26",
            "A27", "A28", "A29", "A30", "A31", "A32", "A33", "A34", "A35", "A36", "A37", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50")
                             , dtype={"A6": str, "A8": str})
        xlsx[['A51', 'A52', 'A53', "A54", "A55", 'A56', 'A57', 'A58', 'A59', 'A60', "A61", "A62", 'A63', 'A64', 'A65',
              'A66']] = pd.DataFrame(
            [[None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None]],
            index=xlsx.index)
        xlsx = xlsx[
            ['A1', 'A6', 'A8', 'A10', 'A11', 'A25', 'A30', 'A38', 'A39', 'A40', 'A41', 'A42', 'A43', 'A44', 'A45',
             'A46', 'A47', 'A48', 'A49', 'A50', 'A51', 'A52', 'A53', 'A54', 'A55', 'A56', 'A57', 'A58', 'A59', 'A60',
             'A61', 'A62', 'A63', 'A64', 'A65', 'A66']]
        for ind in range(len(xlsx)):
            if type(xlsx.iloc[ind]['A25']) == str:
                xlsx.at[ind, "A63"] = 0

        # ОСНОВНАЯ ЧАСТЬ

        dtb = pd.DataFrame()
        print("1 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["A6"]
            query = (f"""
                SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                      Pers.Phone, Address.Addr, Pers.IDPers, Address.House, Address.Stroenie,Address.Corp,Address.Flat
                FROM Pers INNER JOIN
                Polis ON Pers.IDPers = Polis.IDPers INNER JOIN
                Address ON Pers.IDPers = Address.IDAddressOwner
                WHERE (Polis.PolisDateF IS NULL) and Pers.ENP = '{id_excel}'
                ORDER BY  address.IDAddressType Desc
                """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        # xlsx.to_excel("t1.xlsx")

        print("2 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A6'] == xlsx.iloc[ind]['ENP'] and xlsx.iloc[ind1]['A63'] != 0:
                    # adddr = "д." + xlsx.iloc[ind]['House']
                    # adddr = adddr.replace("-", ",кв.")
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A56'] = xlsx.iloc[ind]['Addr']  # + " " + adddr
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    # xlsx.at[ind1, 'A66'] = xlsx.iloc[ind]['House']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'Addr', 'IDPers'], axis=1,
                  inplace=True)
        xlsx.to_excel("t1.xlsx")

        # С БЕЗДОМНЫМИ
        dtb = pd.DataFrame()
        print("3 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["A6"]
            query = (f"""
                               SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                Pers.Phone, Pers.IDPers
                FROM Pers INNER JOIN
                Polis ON Pers.IDPers = Polis.IDPers
                WHERE (Polis.PolisDateF IS NULL) and Pers.ENP = '{id_excel}'
                """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        # xlsx.to_excel("concat.xlsx")

        print("4 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A6'] == xlsx.iloc[ind]['ENP'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'IDPers'], axis=1, inplace=True)
        # xlsx.to_excel("t2.xlsx")
        # по временному
        dtb = pd.DataFrame()
        print("5 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["A8"]
            query = (f"""
                               SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                Pers.Phone, Address.Addr, Polis.PolisN, Pers.IDPers
                FROM Pers INNER JOIN
                Polis ON Pers.IDPers = Polis.IDPers INNER JOIN
                Address ON Pers.IDPers = Address.IDAddressOwner
                WHERE (Polis.PolisDateF IS NULL) and Polis.PolisN= '{id_excel}'
                ORDER BY  address.IDAddressType Desc
                """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        xlsx.to_excel("concat.xlsx")
        print("6 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A8'] == xlsx.iloc[ind]['PolisN'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A56'] = xlsx.iloc[ind]['Addr']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'Addr', 'PolisN', 'IDPers'], axis=1,
                  inplace=True)
        # xlsx.to_excel("t3.xlsx")
        dtb = pd.DataFrame()
        # недействующие
        print("7 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["A6"]

            query = (f"""
                              SELECT top 1 Pers.ENP,  Polis.PolisDateF,  _sCloseStatus.Name, Pers.IDPers
                FROM Pers INNER JOIN
                Polis ON Pers.IDPers = Polis.IDPers inner join
                _sCloseStatus on _sCloseStatus.IDCloseStatus = Polis.IDCloseStatus
                WHERE (Polis.PolisDateF IS not NULL) and Pers.Enp = '{id_excel}'
                ORDER BY  Polis.PolisDateF Desc
                """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        xlsx = pd.concat([xlsx, dtb], axis=1)
        xlsx.to_excel("concat1.xlsx")

        print("8 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A6'] == xlsx.iloc[ind]['ENP'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A57'] = xlsx.iloc[ind]['PolisDateF']
                    xlsx.at[ind1, 'A58'] = xlsx.iloc[ind]['Name']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    xlsx.at[ind1, "A63"] = 0
                    xlsx.loc[ind, 'A11'] = None
                    break
        xlsx.drop(columns=['ENP', 'PolisDateF', 'Name', 'IDPers'], axis=1, inplace=True)
        xlsx.to_excel("got1.xlsx")
        # выбор номеров
        c = []
        print("9 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            if type(xlsx.iloc[ind]['A58']) is str or xlsx.iloc[ind]['A63'] == 0:
                continue
            elif type(xlsx.iloc[ind]['A11']) is not str and type(xlsx.iloc[ind]['A55']) is not str:
                xlsx.loc[ind, 'A60'] = None
            elif type(xlsx.iloc[ind]['A11']) is str:
                b = ''.join(xlsx.iloc[ind]['A11'].split())
                b = ''.join(b.split("+"))
                b = ''.join(b.split("("))
                b = ''.join(b.split(")"))
                b = ''.join(b.split("-"))
                match = re.findall(r'[78][9]\d{9}', b)
                rematch = re.findall(r'\d{6}', b)
                if match != c:
                    xlsx.at[ind, 'A60'] = "смс"
                    xlsx.at[ind, 'A61'] = match
                elif type(xlsx.iloc[ind]['A55']) is str:
                    a = ''.join(xlsx.iloc[ind]['A55'].split())
                    a = ''.join(a.split("+"))
                    a = ''.join(a.split("("))
                    a = ''.join(a.split(")"))
                    a = ''.join(a.split("-"))
                    a = ''.join(itertools.filterfalse(str.isalpha, a))
                    sot = re.findall(r'[9]\d{9}', a)
                    zvon = re.findall(r'd{6}?', a)
                    if sot != c:
                        xlsx.at[ind, 'A60'] = "смс"
                        xlsx.at[ind, 'A61'] = sot[0]
                    elif zvon != c:
                        xlsx.at[ind, 'A60'] = "звонок"
                        xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A55']
                    elif rematch != c:
                        xlsx.at[ind, 'A60'] = "звонок"
                        xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A11']

                    else:
                        xlsx.at[ind, 'A60'] = "письмо"
                else:
                    xlsx.at[ind, 'A60'] = "письмо"
            elif type(xlsx.iloc[ind]['A55']) is str:
                e = ''.join(xlsx.iloc[ind]['A55'].split())
                e = ''.join(e.split("+"))
                e = ''.join(e.split("("))
                e = ''.join(e.split(")"))
                e = ''.join(e.split("-"))
                e = ''.join(itertools.filterfalse(str.isalpha, e))
                sot = re.findall(r'[9]\d{9}', e)
                zvon = re.findall(r'\d{6}?', e)
                if sot != c:
                    xlsx.at[ind, 'A60'] = "смс"
                    xlsx.at[ind, 'A61'] = sot[0]
                elif zvon != c:
                    xlsx.at[ind, 'A60'] = "звонок"
                    xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A55']
                else:
                    xlsx.at[ind, 'A60'] = "письмо"
            else:
                xlsx.at[ind, 'A60'] = "письмо"
        for ind in tqdm(range(len(xlsx))):
            a = str(xlsx.iloc[ind]['A61'])
            a = ''.join(a.split())
            a = ''.join(a.split("["))
            a = ''.join(a.split("]"))
            a = ''.join(a.split("'"))
            a = ''.join(a.split("+"))
            a = ''.join(a.split("("))
            a = ''.join(a.split(")"))
            a = ''.join(a.split("-"))
            a = ''.join(itertools.filterfalse(str.isalpha, a))
            xlsx.at[ind, 'A61'] = a
        # повторные
        xlsx.to_excel("mod1.xlsx")
        xlsx = pd.read_excel("mod1.xlsx", dtype={"A6": str, "A8": str})

        dtb = pd.DataFrame()
        print("10 этап из 12:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["A6"]
            query = (f"""
                                SELECT top 1 A6 as enp, coalesce(A39, A40, A41, A42) as [date], case
                 when A39 is null and A41 is null and A42 is null and A40 is not null then 2
                 when A39 is null and A40 is null and A42 is null and A41 is not null then 6
                 when A40 is null and A41 is null and A39 is null and A42 is not null then 1
                 when A40 is null and A41 is null and A42 is null and A39 is not null then 4
                 when A40 is null and A41 is null and A39 is null and A42 is null then 'null'
                 end as [send by]
                 from informed
                 where coalesce(A39, A40, A41, A42) >= '20220101' and A6 = '{id_excel}'
                 order by [date] desc, [send by] desc
                 """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        # print(xlsx)
        xlsx.to_excel("concat.xlsx")

        print("11 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A63'] != 0 and xlsx.iloc[ind1]['A63'] != 31 and type(xlsx.iloc[ind1]['A51']) != str:
                    xlsx.at[ind1, 'A63'] = 0
                    xlsx.at[ind1, 'A60'] = None
                elif xlsx.iloc[ind1]['A6'] == xlsx.iloc[ind]['enp'] and xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A62'] = xlsx.iloc[ind]['date']
                    xlsx.at[ind1, 'A64'] = xlsx.iloc[ind]['send by']
                    xlsx.at[ind1, 'A63'] = 31
                    continue
        for ind in tqdm(range(len(xlsx))):

            if xlsx.iloc[ind]['A63'] != 31 and xlsx.iloc[ind]['A63'] != 0:
                xlsx.at[ind, 'A63'] = 29

        xlsx.drop(columns=['enp', 'date', 'send by'], axis=1, inplace=True)
        xlsx.to_excel("done12.xlsx")
        inf = pd.read_excel(dirname + "/Справочники/Sprav.xlsx", dtype={"To": str})
        # inf.drop(columns=['Account Name','Traffic Source', 'Communication Name','Communication Scheduled For', 'Communication Template','From','Message Id','Country Prefix','Country Name','Network Name','Purchase Price','Reason','Action', 'Error Group','Error Name', 'Done At', 'Text','Messages Count','Service Name','User Name','Paired Message Id','Clicks', 'Data Payload', 'Communication Start Date'],axis = 1, inplace=True)
        xlsx = pd.concat([xlsx, inf], axis=1)

        print("12 этап из 12:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A30'] == xlsx.iloc[ind]['name']:
                    xlsx.at[ind1, 'A30'] = xlsx.iloc[ind]['name_new']
                    break
        # xlsx.drop(columns=['To','Send At','Status', 'Seen At'],axis = 1, inplace=True)
        xlsx.drop(columns=['Unnamed: 0', 'name', 'name_new'], axis=1, inplace=True)
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)", 'A38': "Способ информирования_эл.почта",
                             'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)", 'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_первичная обработка.xlsx')
        print(datetime.now() - start_time)

    if event == 'let29':
        dt = datetime.now()
        rabname = os.path.dirname(r"z:\данные из ГИС ОМС\УД\Письма/")
        dirname = os.path.dirname(__file__)
        if os.path.exists(rabname):
            shutil.rmtree(rabname, ignore_errors=True)
        os.mkdir(rabname)
        os.mkdir(rabname + "/Penza")
        os.mkdir(rabname + "/bashmakovsky")
        os.mkdir(rabname + "/bekovsky")
        os.mkdir(rabname + "/belinsky")
        os.mkdir(rabname + "/bessonovsky")
        os.mkdir(rabname + "/gorodishensky")
        os.mkdir(rabname + "/zarechny")
        os.mkdir(rabname + "/zemetchisky")
        os.mkdir(rabname + "/issinsky")
        os.mkdir(rabname + "/kamensky")
        os.mkdir(rabname + "/kameshirsky")
        os.mkdir(rabname + "/kolisheisky")
        os.mkdir(rabname + "/kuznetsky")
        os.mkdir(rabname + "/lopatinsky")
        os.mkdir(rabname + "/luninsky")
        os.mkdir(rabname + "/maloserdobsky")
        os.mkdir(rabname + "/mokshansky")
        os.mkdir(rabname + "/narovchatsky")
        os.mkdir(rabname + "/neverkinsky")
        os.mkdir(rabname + "/nizhnelomovsky")
        os.mkdir(rabname + "/nikolsky")
        os.mkdir(rabname + "/pachelmsky")
        os.mkdir(rabname + "/penzensky")
        os.mkdir(rabname + "/serdobsky")
        os.mkdir(rabname + "/sosnovoborsky")
        os.mkdir(rabname + "/spassky")
        os.mkdir(rabname + "/shemisheisky")
        os.mkdir(rabname + "/tamalinsky")
        os.mkdir(rabname + "/vadinsky")
        os.mkdir(rabname + "/other")

        xlsx = pd.read_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_первичная обработка.xlsx', names=(
            "A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str, 'A61': str})
        print('этап 1 из 1:')
        for ind in tqdm(range(len(xlsx))):
            if xlsx.iloc[ind]['A60'] == "письмо" or xlsx.iloc[ind]['A60'] == "звонок" and xlsx.iloc[ind]['A63'] == 29:
                xlsx.at[ind, 'A42'] = dt.strftime("%d/%m/%Y").replace("/", ".")
                xlsx.at[ind, 'A40'] = None
                xlsx.at[ind, 'A39'] = None
                xlsx.at[ind, 'A41'] = None
                xlsx.at[ind, 'A66'] = 1
                if type(xlsx.iloc[ind]['A56']) is str:
                    nme = xlsx.iloc[ind]['A52']
                    if type(xlsx.iloc[ind]['A53']) is str:
                        surname = xlsx.iloc[ind]['A53']
                    elif type(xlsx.iloc[ind]['A53']) is not str:
                        surname = ""
                    address = xlsx.iloc[ind]['A30']

                    addr = xlsx.iloc[ind]['A56']
                    g = re.findall(r"Пензенская", addr)
                    if g == ["Пензенская"]:
                        addr = addr.replace("Пензенская (обл),", "")
                    text_split = addr.split(',')
                    text1 = ','.join(text_split[:2])
                    text2 = ','.join(text_split[2:])
                    # image = Image.open("shabl11.jpg")
                    image1 = Image.open("empty3.jpg")

                    font = ImageFont.truetype("timesbd.ttf", 14)
                    font1 = ImageFont.truetype("times.ttf", 14)
                    # drawer = ImageDraw.Draw(image)
                    drawer1 = ImageDraw.Draw(image1)
                    # drawer.text((890, 820), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font, fill='black')
                    drawer1.text((230, 250), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font1,
                                 fill='black')
                    indzastr = (r"Индивидуальная застройка", addr)
                    # if indzastr != c:
                    #     text1 = ''.join(text_split[:4])
                    #     text2 = ''.join(text_split[4:])
                    if len(address) > 30:
                        # drawer.text((650, 1100), address, font=font, fill='black')
                        drawer1.text((245, 320), address, font=font1, fill='black')
                    elif len(address) <= 30:
                        # drawer.text((900, 1100), address, font=font, fill='black')
                        drawer1.text((215, 320), address, font=font1, fill='black')
                    # drawer.text((295, 924), "Ваша", font=font, fill='black')
                    # drawer.text((1415, 432), addr, font=font, fill='black')
                    # drawer.text((1380, 452), text1, font=font, fill='black')
                    drawer1.text((410, 170), text1, font=font1, fill='black')
                    # drawer.text((1380, 502), text2, font=font, fill='black')
                    drawer1.text((410, 185), text2, font=font1, fill='black')

                    penza = re.findall(r"Пенза", addr, re.IGNORECASE)
                    if penza != c:
                        out = f"{rabname}/Penza/penza{ind}.gif"
                        image1.save(out)  # вот это сохранение зациклить
                        continue
                    # penza1 = re.findall(r"ПЕНЗА", addr)
                    # if penza1 != c:
                    #     image1.save(f"{rabname}/Письма/Penza/penza{ind}.gif")  # вот это сохранение зациклить
                    #     continue
                    bashmakovsky = re.findall(r"Башмаковский ", addr, re.IGNORECASE)
                    if bashmakovsky != c:
                        image1.save(f"{rabname}/bashmakovsky/bashmakovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    bekovsky = re.findall(r"Бековский", addr, re.IGNORECASE)
                    if bekovsky != c:
                        image1.save(f"{rabname}/bekovsky/bekovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    belinsky = re.findall(r"Белинский", addr, re.IGNORECASE)
                    if belinsky != c:
                        image1.save(f"{rabname}/belinsky/belinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    bessonovsky = re.findall(r"Бессоновский", addr, re.IGNORECASE)
                    if bessonovsky != c:
                        image1.save(f"{rabname}/bessonovsky/bessonovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    gorodishensky = re.findall(r"Городищенский", addr, re.IGNORECASE)
                    if gorodishensky != c:
                        image1.save(f"{rabname}/gorodishensky/gorodishensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    zarechny = re.findall(r"Заречный", addr, re.IGNORECASE)
                    if zarechny != c:
                        image1.save(f"{rabname}/zarechny/zarechny{ind}.gif")  # вот это сохранение зациклить
                        continue
                    zemetchisky = re.findall(r"Земетчинский", addr, re.IGNORECASE)
                    if zemetchisky != c:
                        image1.save(f"{rabname}/zemetchisky/zemetchisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    issinsky = re.findall(r"Иссинский", addr, re.IGNORECASE)
                    if issinsky != c:
                        image1.save(f"{rabname}/issinsky/issinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kamensky = re.findall(r"Каменский", addr, re.IGNORECASE)
                    if kamensky != c:
                        image1.save(f"{rabname}/kamensky/kamensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kameshirsky = re.findall(r"Камешкирский", addr, re.IGNORECASE)
                    if kameshirsky != c:
                        image1.save(f"{rabname}/kameshirsky/kameshirsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kolisheisky = re.findall(r"Колышлейский", addr, re.IGNORECASE)
                    if kolisheisky != c:
                        image1.save(f"{rabname}/kolisheisky/kolisheisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kuznetsky = re.findall(r"Кузнецк", addr, re.IGNORECASE)
                    if kuznetsky != c:
                        image1.save(f"{rabname}/kuznetsky/kuznetsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    lopatinsky = re.findall(r"Лопатинский", addr, re.IGNORECASE)
                    if lopatinsky != c:
                        image1.save(f"{rabname}/lopatinsky/lopatinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    luninsky = re.findall(r"Лунинский", addr, re.IGNORECASE)
                    if luninsky != c:
                        image1.save(f"{rabname}/luninsky/luninsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    maloserdobsky = re.findall(r"Малосердобинский", addr, re.IGNORECASE)
                    if maloserdobsky != c:
                        image1.save(f"{rabname}/maloserdobsky/maloserdobsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    mokshansky = re.findall(r"Мокшанский", addr, re.IGNORECASE)
                    if mokshansky != c:
                        image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    mokshansky1 = re.findall(r"МОКШАНСКИЙ", addr, re.IGNORECASE)
                    if mokshansky1 != c:
                        image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    narovchatsky = re.findall(r"Наровчатский", addr, re.IGNORECASE)
                    if narovchatsky != c:
                        image1.save(f"{rabname}/narovchatsky/narovchatsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    neverkinsky = re.findall(r"Неверкинский", addr, re.IGNORECASE)
                    if neverkinsky != c:
                        image1.save(f"{rabname}/neverkinsky/neverkinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    nizhnelomovsky = re.findall(r"Нижнеломовский", addr, re.IGNORECASE)
                    if nizhnelomovsky != c:
                        image1.save(f"{rabname}/nizhnelomovsky/nizhnelomovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    nikolsky = re.findall(r"Никольский", addr, re.IGNORECASE)
                    if nikolsky != c:
                        image1.save(f"{rabname}/nikolsky/nikolsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    pachelmsky = re.findall(r"Пачелмский", addr, re.IGNORECASE)
                    if pachelmsky != c:
                        image1.save(f"{rabname}/pachelmsky/pachelmsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    penzensky = re.findall(r"Пензенский", addr, re.IGNORECASE)
                    if penzensky != c:
                        image1.save(f"{rabname}/penzensky/penzensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    serdobsky = re.findall(r"Сердобский", addr, re.IGNORECASE)
                    if serdobsky != c:
                        image1.save(f"{rabname}/serdobsky/serdobsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    sosnovoborsky = re.findall(r"Сосновоборский", addr, re.IGNORECASE)
                    if sosnovoborsky != c:
                        image1.save(f"{rabname}/sosnovoborsky/sosnovoborsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    spassky = re.findall(r"Спасский", addr, re.IGNORECASE)
                    if spassky != c:
                        image1.save(f"{rabname}/spassky/spassky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    tamalinsky = re.findall(r"Тамалинский", addr, re.IGNORECASE)
                    if tamalinsky != c:
                        image1.save(f"{rabname}/tamalinsky/tamalinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    shemisheisky = re.findall(r"Шемышейский", addr, re.IGNORECASE)
                    if shemisheisky != c:
                        image1.save(f"{rabname}/shemisheisky/shemisheisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    vadinsky = re.findall(r"Вадинский", addr, re.IGNORECASE)
                    if vadinsky != c:
                        image1.save(f"{rabname}/vadinsky/vadinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    else:
                        image1.save(f"{rabname}/other/other{ind}.gif")  # вот это сохранение зациклить
                        continue
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка\SMO_после_писем.xlsx')
    if event == 'let31':
        dt = datetime.now()
        rabname = os.path.dirname(r"z:\данные из ГИС ОМС\УД\Письма/")
        dirname = os.path.dirname(__file__)
        if os.path.exists(rabname):
            shutil.rmtree(rabname, ignore_errors=True)
        os.mkdir(rabname)
        os.mkdir(rabname + "/Penza")
        os.mkdir(rabname + "/bashmakovsky")
        os.mkdir(rabname + "/bekovsky")
        os.mkdir(rabname + "/belinsky")
        os.mkdir(rabname + "/bessonovsky")
        os.mkdir(rabname + "/gorodishensky")
        os.mkdir(rabname + "/zarechny")
        os.mkdir(rabname + "/zemetchisky")
        os.mkdir(rabname + "/issinsky")
        os.mkdir(rabname + "/kamensky")
        os.mkdir(rabname + "/kameshirsky")
        os.mkdir(rabname + "/kolisheisky")
        os.mkdir(rabname + "/kuznetsky")
        os.mkdir(rabname + "/lopatinsky")
        os.mkdir(rabname + "/luninsky")
        os.mkdir(rabname + "/maloserdobsky")
        os.mkdir(rabname + "/mokshansky")
        os.mkdir(rabname + "/narovchatsky")
        os.mkdir(rabname + "/neverkinsky")
        os.mkdir(rabname + "/nizhnelomovsky")
        os.mkdir(rabname + "/nikolsky")
        os.mkdir(rabname + "/pachelmsky")
        os.mkdir(rabname + "/penzensky")
        os.mkdir(rabname + "/serdobsky")
        os.mkdir(rabname + "/sosnovoborsky")
        os.mkdir(rabname + "/spassky")
        os.mkdir(rabname + "/shemisheisky")
        os.mkdir(rabname + "/tamalinsky")
        os.mkdir(rabname + "/vadinsky")
        os.mkdir(rabname + "/other")

        xlsx = pd.read_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_первичная обработка.xlsx', names=(
            "A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str})
        print('этап 1 из 1:')
        for ind in tqdm(range(len(xlsx))):
            if xlsx.iloc[ind]['A60'] == "письмо" or xlsx.iloc[ind]['A60'] == "звонок" and xlsx.iloc[ind]['A63'] == 31:
                xlsx.at[ind, 'A42'] = dt.strftime("%d/%m/%Y").replace("/", ".")
                xlsx.at[ind, 'A40'] = None
                xlsx.at[ind, 'A39'] = None
                xlsx.at[ind, 'A41'] = None
                xlsx.at[ind, 'A66'] = 1
                if type(xlsx.iloc[ind]['A56']) is str:
                    nme = xlsx.iloc[ind]['A52']
                    if type(xlsx.iloc[ind]['A53']) is str:
                        surname = xlsx.iloc[ind]['A53']
                    elif type(xlsx.iloc[ind]['A53']) is not str:
                        surname = ""
                    address = xlsx.iloc[ind]['A30']

                    addr = xlsx.iloc[ind]['A56']
                    g = re.findall(r"Пензенская", addr)
                    if g == ["Пензенская"]:
                        addr = addr.replace("Пензенская (обл),", "")
                    text_split = addr.split(',')
                    text1 = ','.join(text_split[:2])
                    text2 = ','.join(text_split[2:])
                    # image = Image.open("shabl11.jpg")
                    image1 = Image.open("empty3.jpg")

                    font = ImageFont.truetype("timesbd.ttf", 14)
                    font1 = ImageFont.truetype("times.ttf", 14)
                    # drawer = ImageDraw.Draw(image)
                    drawer1 = ImageDraw.Draw(image1)
                    # drawer.text((890, 820), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font, fill='black')
                    drawer1.text((230, 250), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font1,
                                 fill='black')
                    indzastr = (r"Индивидуальная застройка", addr)
                    # if indzastr != c:
                    #     text1 = ''.join(text_split[:4])
                    #     text2 = ''.join(text_split[4:])
                    if len(address) > 30:
                        # drawer.text((650, 1100), address, font=font, fill='black')
                        drawer1.text((245, 320), address, font=font1, fill='black')
                    elif len(address) <= 30:
                        # drawer.text((900, 1100), address, font=font, fill='black')
                        drawer1.text((215, 320), address, font=font1, fill='black')
                    # drawer.text((295, 924), "Ваша", font=font, fill='black')
                    # drawer.text((1415, 432), addr, font=font, fill='black')
                    # drawer.text((1380, 452), text1, font=font, fill='black')
                    drawer1.text((410, 170), text1, font=font1, fill='black')
                    # drawer.text((1380, 502), text2, font=font, fill='black')
                    drawer1.text((410, 185), text2, font=font1, fill='black')

                    penza = re.findall(r"Пенза", addr, re.IGNORECASE)
                    if penza != c:
                        out = f"{rabname}/Penza/penza{ind}.gif"
                        image1.save(out)  # вот это сохранение зациклить
                        continue
                    # penza1 = re.findall(r"ПЕНЗА", addr)
                    # if penza1 != c:
                    #     image1.save(f"{rabname}/Письма/Penza/penza{ind}.gif")  # вот это сохранение зациклить
                    #     continue
                    bashmakovsky = re.findall(r"Башмаковский ", addr, re.IGNORECASE)
                    if bashmakovsky != c:
                        image1.save(f"{rabname}/bashmakovsky/bashmakovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    bekovsky = re.findall(r"Бековский", addr, re.IGNORECASE)
                    if bekovsky != c:
                        image1.save(f"{rabname}/bekovsky/bekovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    belinsky = re.findall(r"Белинский", addr, re.IGNORECASE)
                    if belinsky != c:
                        image1.save(f"{rabname}/belinsky/belinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    bessonovsky = re.findall(r"Бессоновский", addr, re.IGNORECASE)
                    if bessonovsky != c:
                        image1.save(f"{rabname}/bessonovsky/bessonovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    gorodishensky = re.findall(r"Городищенский", addr, re.IGNORECASE)
                    if gorodishensky != c:
                        image1.save(f"{rabname}/gorodishensky/gorodishensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    zarechny = re.findall(r"Заречный", addr, re.IGNORECASE)
                    if zarechny != c:
                        image1.save(f"{rabname}/zarechny/zarechny{ind}.gif")  # вот это сохранение зациклить
                        continue
                    zemetchisky = re.findall(r"Земетчинский", addr, re.IGNORECASE)
                    if zemetchisky != c:
                        image1.save(f"{rabname}/zemetchisky/zemetchisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    issinsky = re.findall(r"Иссинский", addr, re.IGNORECASE)
                    if issinsky != c:
                        image1.save(f"{rabname}/issinsky/issinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kamensky = re.findall(r"Каменский", addr, re.IGNORECASE)
                    if kamensky != c:
                        image1.save(f"{rabname}/kamensky/kamensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kameshirsky = re.findall(r"Камешкирский", addr, re.IGNORECASE)
                    if kameshirsky != c:
                        image1.save(f"{rabname}/kameshirsky/kameshirsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kolisheisky = re.findall(r"Колышлейский", addr, re.IGNORECASE)
                    if kolisheisky != c:
                        image1.save(f"{rabname}/kolisheisky/kolisheisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    kuznetsky = re.findall(r"Кузнецк", addr, re.IGNORECASE)
                    if kuznetsky != c:
                        image1.save(f"{rabname}/kuznetsky/kuznetsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    lopatinsky = re.findall(r"Лопатинский", addr, re.IGNORECASE)
                    if lopatinsky != c:
                        image1.save(f"{rabname}/lopatinsky/lopatinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    luninsky = re.findall(r"Лунинский", addr, re.IGNORECASE)
                    if luninsky != c:
                        image1.save(f"{rabname}/luninsky/luninsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    maloserdobsky = re.findall(r"Малосердобинский", addr, re.IGNORECASE)
                    if maloserdobsky != c:
                        image1.save(f"{rabname}/maloserdobsky/maloserdobsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    mokshansky = re.findall(r"Мокшанский", addr, re.IGNORECASE)
                    if mokshansky != c:
                        image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    mokshansky1 = re.findall(r"МОКШАНСКИЙ", addr, re.IGNORECASE)
                    if mokshansky1 != c:
                        image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    narovchatsky = re.findall(r"Наровчатский", addr, re.IGNORECASE)
                    if narovchatsky != c:
                        image1.save(f"{rabname}/narovchatsky/narovchatsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    neverkinsky = re.findall(r"Неверкинский", addr, re.IGNORECASE)
                    if neverkinsky != c:
                        image1.save(f"{rabname}/neverkinsky/neverkinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    nizhnelomovsky = re.findall(r"Нижнеломовский", addr, re.IGNORECASE)
                    if nizhnelomovsky != c:
                        image1.save(f"{rabname}/nizhnelomovsky/nizhnelomovsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    nikolsky = re.findall(r"Никольский", addr, re.IGNORECASE)
                    if nikolsky != c:
                        image1.save(f"{rabname}/nikolsky/nikolsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    pachelmsky = re.findall(r"Пачелмский", addr, re.IGNORECASE)
                    if pachelmsky != c:
                        image1.save(f"{rabname}/pachelmsky/pachelmsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    penzensky = re.findall(r"Пензенский", addr, re.IGNORECASE)
                    if penzensky != c:
                        image1.save(f"{rabname}/penzensky/penzensky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    serdobsky = re.findall(r"Сердобский", addr, re.IGNORECASE)
                    if serdobsky != c:
                        image1.save(f"{rabname}/serdobsky/serdobsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    sosnovoborsky = re.findall(r"Сосновоборский", addr, re.IGNORECASE)
                    if sosnovoborsky != c:
                        image1.save(f"{rabname}/sosnovoborsky/sosnovoborsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    spassky = re.findall(r"Спасский", addr, re.IGNORECASE)
                    if spassky != c:
                        image1.save(f"{rabname}/spassky/spassky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    tamalinsky = re.findall(r"Тамалинский", addr, re.IGNORECASE)
                    if tamalinsky != c:
                        image1.save(f"{rabname}/tamalinsky/tamalinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    shemisheisky = re.findall(r"Шемышейский", addr, re.IGNORECASE)
                    if shemisheisky != c:
                        image1.save(f"{rabname}/shemisheisky/shemisheisky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    vadinsky = re.findall(r"Вадинский", addr, re.IGNORECASE)
                    if vadinsky != c:
                        image1.save(f"{rabname}/vadinsky/vadinsky{ind}.gif")  # вот это сохранение зациклить
                        continue
                    else:
                        image1.save(f"{rabname}/other/other{ind}.gif")  # вот это сохранение зациклить
                        continue
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_после_писем.xlsx')
    if event == 'Viber':
        import pandas as pd

        xlsx = pd.read_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_после_писем.xlsx', names=(
            "A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str, "A61": str})
        pathh = values['-FILE_PATH_VIBER-']
        inf = pd.read_excel(pathh, dtype={"To": str})
        inf = inf[['To', 'Send At']]
        xlsx = pd.concat([xlsx, inf], axis=1)

        print("1 этап из 2:")
        for ind in tqdm(range(len(xlsx))):
            if type(xlsx.iloc[ind]['A61']) is str:
                if len(xlsx.iloc[ind]['A61']) == 10:
                    xlsx.at[ind, 'A61'] = "7" + xlsx.iloc[ind]['A61']
                # if xlsx.iloc[ind]['A61'] ==  re.findall(r'[78][9]\d{9}', xlsx.iloc[ind]['A61']) and xlsx.iloc[ind, 'A60'] == "смс" :
                #     xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A61'][1:]
        xlsx.to_excel("out_v.xlsx")
        print("2 этап из 2:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A60'] == "смс" and xlsx.iloc[ind1]['A63'] != 0:
                    if xlsx.iloc[ind1]['A61'] == xlsx.iloc[ind]['To']:
                        xlsx.at[ind1, 'A41'] = xlsx.iloc[ind]['Send At'].replace("/", ".").split(" ")[0]
                        xlsx.at[ind1, 'A40'] = None
                        xlsx.at[ind1, 'A39'] = None
                        xlsx.at[ind1, 'A42'] = None
                        # xlsx.at[ind1, 'A65'] = xlsx.iloc[ind]['Status']
                        xlsx.at[ind1, 'A66'] = 6
                        # if xlsx.iloc[ind1]['A63'] != 33 and xlsx.iloc[ind1]['A63'] != 31:
                        # xlsx.at[ind1, 'A63'] = 29

        xlsx.drop(columns=['To', 'Send At'], axis=1, inplace=True)
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка/out_viber.xlsx')
    if event == 'Sms':
        xlsx = pd.read_excel('Z:/Данные из ГИС ОМС/УД/Обработка/out_viber.xlsx', names=(
            "A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str, "A61": str})
        pathh = values['-FILE_PATH_SMS-']
        inf = pd.read_excel(pathh, dtype={"To": str})
        inf = inf[['To', 'Send At']]
        xlsx = pd.concat([xlsx, inf], axis=1)

        # print("1 этап из 2:")
        # for ind in tqdm(range(len(xlsx))):
        #     if type(xlsx.iloc[ind]['A61']) is str:
        #         if len(xlsx.iloc[ind]['A61']) == 10:
        #             xlsx.at[ind, 'A61'] = "7" + xlsx.iloc[ind]['A61']
        # xlsx.to_excel("out.xlsx")
        print("1 этап из 2:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A60'] == "смс" and xlsx.iloc[ind1]['A63'] != 0:
                    if xlsx.iloc[ind1]['A61'] == xlsx.iloc[ind]['To']:
                        xlsx.at[ind1, 'A40'] = xlsx.iloc[ind]['Send At'].replace("/", ".").split(" ")[0]
                        xlsx.at[ind1, 'A41'] = None
                        xlsx.at[ind1, 'A39'] = None
                        xlsx.at[ind1, 'A42'] = None
                        #xlsx.at[ind1, 'A65'] = xlsx.iloc[ind]['Status']
                        xlsx.at[ind1, 'A66'] = 2
                        # if xlsx.iloc[ind1]['A63'] != 33 and xlsx.iloc[ind1]['A63'] != 31:
                        #     xlsx.at[ind1, 'A63'] = 29
                        break
        xlsx.drop(columns=['To', 'Send At'], axis=1, inplace=True)
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/Обработка/SMO_out_sms.xlsx')
    if event == 'Base':
        server = "10.58.1.200"
        dbname = "RGSNEW1"
        uname = "rgs"
        pword = "233239"
        pathh = values['Base']
        xlsx = pd.read_excel(pathh, names=(
            "A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str, "A61": str})
        xlsx = xlsx[["A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
                     "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50"]]

        # eng = create_engine('mssql+pyodbc://DMITRY-PC\SQLEXPRESS/RGSNEW1?driver=SQL+Server')
        eng = create_engine(
            "mssql+pyodbc://" + uname + ":" + pword + "@" + server + "/" + dbname + "?driver=SQL+Server")
        xlsx.to_sql("informed", eng, if_exists="append", index=False)
    if event == 'idisp':
        pathh = values["-FILE_PATH_INDEXDISP-"]
        xlsx = pd.read_excel(pathh,  header=4, names=(
            "Index", "A2", "A3", "A4", "A5", "A90", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15",
            "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26",
            "A27", "A28", "A29", "A30", "A31", "A32", "A33", "A34", "A35", "A36", "A37", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50")
                            , dtype={"A90": str})
        dtb = pd.read_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_прошлый.xlsx', names=("A1", "A6", "A8", "A10", "A11", "A25", "A30", "A38", "A39",
            "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", "A48", "A49", "A50", "A51", "A52", "A53", "A54",
            "A55", "A56", "A57", "A58", "A59", "A60",
            "A61", "A62", "A63", "A64", "A65", "A66", 'House', 'Stroenie', 'Corp', 'Flat'),
                             dtype={"A6": str, "A8": str, "A61": str})
        xlsx = xlsx[['Index', 'A90']]
        dtb.to_excel(dirname + "/Выходные файлы/SMO_changed1.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        print("1 этап из 1:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['A6'] == xlsx.iloc[ind]['A90']:
                    xlsx.at[ind1, 'A1'] = xlsx.iloc[ind]['Index']
                    break
        # xlsx.drop(columns=['To','Send At','Status', 'Seen At'],axis = 1, inplace=True)
        xlsx.drop(columns=['Index', 'A90'], axis=1, inplace=True)
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_измененный_порядок.xlsx')
    if event == 'prof':
        rabname = os.path.dirname(r"z:\данные из ГИС ОМС\УД\Письма/")
        dirname = os.path.dirname(__file__)
        if os.path.exists(rabname):
            shutil.rmtree(rabname, ignore_errors=True)
        os.mkdir(rabname)
        os.mkdir(rabname + "/Penza")
        os.mkdir(rabname + "/bashmakovsky")
        os.mkdir(rabname + "/bekovsky")
        os.mkdir(rabname + "/belinsky")
        os.mkdir(rabname + "/bessonovsky")
        os.mkdir(rabname + "/gorodishensky")
        os.mkdir(rabname + "/zarechny")
        os.mkdir(rabname + "/zemetchisky")
        os.mkdir(rabname + "/issinsky")
        os.mkdir(rabname + "/kamensky")
        os.mkdir(rabname + "/kameshirsky")
        os.mkdir(rabname + "/kolisheisky")
        os.mkdir(rabname + "/kuznetsky")
        os.mkdir(rabname + "/lopatinsky")
        os.mkdir(rabname + "/luninsky")
        os.mkdir(rabname + "/maloserdobsky")
        os.mkdir(rabname + "/mokshansky")
        os.mkdir(rabname + "/narovchatsky")
        os.mkdir(rabname + "/neverkinsky")
        os.mkdir(rabname + "/nizhnelomovsky")
        os.mkdir(rabname + "/nikolsky")
        os.mkdir(rabname + "/pachelmsky")
        os.mkdir(rabname + "/penzensky")
        os.mkdir(rabname + "/serdobsky")
        os.mkdir(rabname + "/sosnovoborsky")
        os.mkdir(rabname + "/spassky")
        os.mkdir(rabname + "/shemisheisky")
        os.mkdir(rabname + "/tamalinsky")
        os.mkdir(rabname + "/vadinsky")
        os.mkdir(rabname + "/other")

        pathh = values["-FILE_PATH_PROF-"]
        xlsx = pd.read_excel(pathh)
        xlsx = xlsx[['im', 'ot', 'namePodr', 'adr']]
        c = []
        print("1 этап из 1:")
        for ind in tqdm(range(len(xlsx))):
            nme = xlsx.iloc[ind]['im']
            if type(xlsx.iloc[ind]['ot']) is str:
                surname = xlsx.iloc[ind]['ot']
            elif type(xlsx.iloc[ind]['ot']) is not str:
                surname = ""
            address = xlsx.iloc[ind]['namePodr']

            addr = xlsx.iloc[ind]['adr']
            g = re.findall(r"Пензенская", addr)
            if g == ["Пензенская"]:
                addr = addr.replace("Пензенская (обл),", "")
            text_split = addr.split(',')
            text1 = ','.join(text_split[:2])
            text2 = ','.join(text_split[2:])
            # image = Image.open("shabl11.jpg")
            image1 = Image.open("empty3.jpg")

            font = ImageFont.truetype("timesbd.ttf", 14)
            font1 = ImageFont.truetype("times.ttf", 14)
            # drawer = ImageDraw.Draw(image)
            drawer1 = ImageDraw.Draw(image1)
            # drawer.text((890, 820), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font, fill='black')
            drawer1.text((230, 250), "Уважаемый(-ая)" + " " + nme + " " + surname + "!", font=font1,
                         fill='black')
            indzastr = (r"Индивидуальная застройка", addr)
            # if indzastr != c:
            #     text1 = ''.join(text_split[:4])
            #     text2 = ''.join(text_split[4:])
            if len(address) > 30:
                # drawer.text((650, 1100), address, font=font, fill='black')
                drawer1.text((245, 320), address, font=font1, fill='black')
            elif len(address) <= 30:
                # drawer.text((900, 1100), address, font=font, fill='black')
                drawer1.text((215, 320), address, font=font1, fill='black')
            # drawer.text((295, 924), "Ваша", font=font, fill='black')
            # drawer.text((1415, 432), addr, font=font, fill='black')
            # drawer.text((1380, 452), text1, font=font, fill='black')
            drawer1.text((410, 170), text1, font=font1, fill='black')
            # drawer.text((1380, 502), text2, font=font, fill='black')
            drawer1.text((410, 185), text2, font=font1, fill='black')

            penza = re.findall(r"Пенза", addr, re.IGNORECASE)
            if penza != c:
                out = f"{rabname}/Penza/penza{ind}.gif"
                image1.save(out)  # вот это сохранение зациклить
                continue
            # penza1 = re.findall(r"ПЕНЗА", addr)
            # if penza1 != c:
            #     image1.save(f"{rabname}/Письма/Penza/penza{ind}.gif")  # вот это сохранение зациклить
            #     continue
            bashmakovsky = re.findall(r"Башмаковский ", addr, re.IGNORECASE)
            if bashmakovsky != c:
                image1.save(f"{rabname}/bashmakovsky/bashmakovsky{ind}.gif")  # вот это сохранение зациклить
                continue
            bekovsky = re.findall(r"Бековский", addr, re.IGNORECASE)
            if bekovsky != c:
                image1.save(f"{rabname}/bekovsky/bekovsky{ind}.gif")  # вот это сохранение зациклить
                continue
            belinsky = re.findall(r"Белинский", addr, re.IGNORECASE)
            if belinsky != c:
                image1.save(f"{rabname}/belinsky/belinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            bessonovsky = re.findall(r"Бессоновский", addr, re.IGNORECASE)
            if bessonovsky != c:
                image1.save(f"{rabname}/bessonovsky/bessonovsky{ind}.gif")  # вот это сохранение зациклить
                continue
            gorodishensky = re.findall(r"Городищенский", addr, re.IGNORECASE)
            if gorodishensky != c:
                image1.save(f"{rabname}/gorodishensky/gorodishensky{ind}.gif")  # вот это сохранение зациклить
                continue
            zarechny = re.findall(r"Заречный", addr, re.IGNORECASE)
            if zarechny != c:
                image1.save(f"{rabname}/zarechny/zarechny{ind}.gif")  # вот это сохранение зациклить
                continue
            zemetchisky = re.findall(r"Земетчинский", addr, re.IGNORECASE)
            if zemetchisky != c:
                image1.save(f"{rabname}/zemetchisky/zemetchisky{ind}.gif")  # вот это сохранение зациклить
                continue
            issinsky = re.findall(r"Иссинский", addr, re.IGNORECASE)
            if issinsky != c:
                image1.save(f"{rabname}/issinsky/issinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            kamensky = re.findall(r"Каменский", addr, re.IGNORECASE)
            if kamensky != c:
                image1.save(f"{rabname}/kamensky/kamensky{ind}.gif")  # вот это сохранение зациклить
                continue
            kameshirsky = re.findall(r"Камешкирский", addr, re.IGNORECASE)
            if kameshirsky != c:
                image1.save(f"{rabname}/kameshirsky/kameshirsky{ind}.gif")  # вот это сохранение зациклить
                continue
            kolisheisky = re.findall(r"Колышлейский", addr, re.IGNORECASE)
            if kolisheisky != c:
                image1.save(f"{rabname}/kolisheisky/kolisheisky{ind}.gif")  # вот это сохранение зациклить
                continue
            kuznetsky = re.findall(r"Кузнецк", addr, re.IGNORECASE)
            if kuznetsky != c:
                image1.save(f"{rabname}/kuznetsky/kuznetsky{ind}.gif")  # вот это сохранение зациклить
                continue
            lopatinsky = re.findall(r"Лопатинский", addr, re.IGNORECASE)
            if lopatinsky != c:
                image1.save(f"{rabname}/lopatinsky/lopatinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            luninsky = re.findall(r"Лунинский", addr, re.IGNORECASE)
            if luninsky != c:
                image1.save(f"{rabname}/luninsky/luninsky{ind}.gif")  # вот это сохранение зациклить
                continue
            maloserdobsky = re.findall(r"Малосердобинский", addr, re.IGNORECASE)
            if maloserdobsky != c:
                image1.save(f"{rabname}/maloserdobsky/maloserdobsky{ind}.gif")  # вот это сохранение зациклить
                continue
            mokshansky = re.findall(r"Мокшанский", addr, re.IGNORECASE)
            if mokshansky != c:
                image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                continue
            mokshansky1 = re.findall(r"МОКШАНСКИЙ", addr, re.IGNORECASE)
            if mokshansky1 != c:
                image1.save(f"{rabname}/mokshansky/mokshansky{ind}.gif")  # вот это сохранение зациклить
                continue
            narovchatsky = re.findall(r"Наровчатский", addr, re.IGNORECASE)
            if narovchatsky != c:
                image1.save(f"{rabname}/narovchatsky/narovchatsky{ind}.gif")  # вот это сохранение зациклить
                continue
            neverkinsky = re.findall(r"Неверкинский", addr, re.IGNORECASE)
            if neverkinsky != c:
                image1.save(f"{rabname}/neverkinsky/neverkinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            nizhnelomovsky = re.findall(r"Нижнеломовский", addr, re.IGNORECASE)
            if nizhnelomovsky != c:
                image1.save(f"{rabname}/nizhnelomovsky/nizhnelomovsky{ind}.gif")  # вот это сохранение зациклить
                continue
            nikolsky = re.findall(r"Никольский", addr, re.IGNORECASE)
            if nikolsky != c:
                image1.save(f"{rabname}/nikolsky/nikolsky{ind}.gif")  # вот это сохранение зациклить
                continue
            pachelmsky = re.findall(r"Пачелмский", addr, re.IGNORECASE)
            if pachelmsky != c:
                image1.save(f"{rabname}/pachelmsky/pachelmsky{ind}.gif")  # вот это сохранение зациклить
                continue
            penzensky = re.findall(r"Пензенский", addr, re.IGNORECASE)
            if penzensky != c:
                image1.save(f"{rabname}/penzensky/penzensky{ind}.gif")  # вот это сохранение зациклить
                continue
            serdobsky = re.findall(r"Сердобский", addr, re.IGNORECASE)
            if serdobsky != c:
                image1.save(f"{rabname}/serdobsky/serdobsky{ind}.gif")  # вот это сохранение зациклить
                continue
            sosnovoborsky = re.findall(r"Сосновоборский", addr, re.IGNORECASE)
            if sosnovoborsky != c:
                image1.save(f"{rabname}/sosnovoborsky/sosnovoborsky{ind}.gif")  # вот это сохранение зациклить
                continue
            spassky = re.findall(r"Спасский", addr, re.IGNORECASE)
            if spassky != c:
                image1.save(f"{rabname}/spassky/spassky{ind}.gif")  # вот это сохранение зациклить
                continue
            tamalinsky = re.findall(r"Тамалинский", addr, re.IGNORECASE)
            if tamalinsky != c:
                image1.save(f"{rabname}/tamalinsky/tamalinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            shemisheisky = re.findall(r"Шемышейский", addr, re.IGNORECASE)
            if shemisheisky != c:
                image1.save(f"{rabname}/shemisheisky/shemisheisky{ind}.gif")  # вот это сохранение зациклить
                continue
            vadinsky = re.findall(r"Вадинский", addr, re.IGNORECASE)
            if vadinsky != c:
                image1.save(f"{rabname}/vadinsky/vadinsky{ind}.gif")  # вот это сохранение зациклить
                continue
            else:
                image1.save(f"{rabname}/other/other{ind}.gif")  # вот это сохранение зациклить
                continue
    if event == 'vact':
        start_time = datetime.now()

        # server = "DMITRY\SQLEXPRESS"
        server = "10.58.1.200"
        dbname = "RGSNEW1"
        uname = "rgs"
        pword = "233239"
        # eng = create_engine("mssql+pyodbc://"+server+"/"+dbname+"?driver=SQL+Server")
        eng = create_engine(
            "mssql+pyodbc://" + uname + ":" + pword + "@" + server + "/" + dbname + "?driver=SQL+Server")
        pathh = values['-FILE_PATH_VACT-']
        # xlsx = pd.read_excel(pathh, dtype={"A6": str, "A8": str})
        xlsx = pd.read_excel(pathh, header=4,
                             dtype={"6": str, "8": str})
        xlsx[['A51', 'A52', 'A53', "A54", "A55", 'A56', 'A57', 'A58', 'A59', 'A60', "A61", "A62", 'A63', 'A64', 'A65',
              'A66']] = pd.DataFrame(
            [[None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None]],
            index=xlsx.index)

        dtb = pd.DataFrame()
        print("1 этап из 9:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["6"]
            query = (f"""
                   SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                         Pers.Phone, Address.Addr, Pers.IDPers, Address.House, Address.Stroenie,Address.Corp,Address.Flat
                   FROM Pers INNER JOIN
                   Polis ON Pers.IDPers = Polis.IDPers INNER JOIN
                   Address ON Pers.IDPers = Address.IDAddressOwner
                   WHERE (Polis.PolisDateF IS NULL) and Pers.ENP = '{id_excel}'
                   ORDER BY  address.IDAddressType Desc
                   """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        xlsx.to_excel("t1.xlsx")

        print("2 этап из 9:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['6'] == xlsx.iloc[ind]['ENP'] and xlsx.iloc[ind1]['A63'] != 0:
                    # adddr = "д." + xlsx.iloc[ind]['House']
                    # adddr = adddr.replace("-", ",кв.")
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A56'] = xlsx.iloc[ind]['Addr']  # + " " + adddr
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    # xlsx.at[ind1, 'A66'] = xlsx.iloc[ind]['House']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'Addr', 'IDPers'], axis=1,
                  inplace=True)
        xlsx.to_excel("t2.xlsx")

        # С БЕЗДОМНЫМИ
        dtb = pd.DataFrame()
        print("3 этап из 9:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["6"]
            query = (f"""
                                  SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                   Pers.Phone, Pers.IDPers
                   FROM Pers INNER JOIN
                   Polis ON Pers.IDPers = Polis.IDPers
                   WHERE (Polis.PolisDateF IS NULL) and Pers.ENP = '{id_excel}'
                   """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        dtb.to_excel("not_concat33.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        xlsx.to_excel("concat33.xlsx")

        print("4 этап из 9:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['6'] == xlsx.iloc[ind]['ENP'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'IDPers'], axis=1, inplace=True)
        xlsx.to_excel("t4.xlsx")
        # по временному
        dtb = pd.DataFrame()
        print("5 этап из 9:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["8"]
            query = (f"""
                                  SELECT top 1 Pers.Surname, Pers.Name1, Pers.Name2, Pers.Birthday, Pers.ENP,
                   Pers.Phone, Address.Addr, Polis.PolisN, Pers.IDPers
                   FROM Pers INNER JOIN
                   Polis ON Pers.IDPers = Polis.IDPers INNER JOIN
                   Address ON Pers.IDPers = Address.IDAddressOwner
                   WHERE (Polis.PolisDateF IS NULL) and Polis.PolisN= '{id_excel}'
                   ORDER BY  address.IDAddressType Desc
                   """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        # dtb.to_excel("not_concat.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        # xlsx.to_excel("concat.xlsx")
        print("6 этап из 9:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['8'] == xlsx.iloc[ind]['PolisN'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A51'] = xlsx.iloc[ind]['Surname']
                    xlsx.at[ind1, 'A52'] = xlsx.iloc[ind]['Name1']
                    xlsx.at[ind1, 'A53'] = xlsx.iloc[ind]['Name2']
                    xlsx.at[ind1, 'A54'] = xlsx.iloc[ind]['Birthday']
                    xlsx.at[ind1, 'A55'] = xlsx.iloc[ind]['Phone']
                    xlsx.at[ind1, 'A56'] = xlsx.iloc[ind]['Addr']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    break
        xlsx.drop(columns=['ENP', 'Surname', 'Name1', 'Name2', 'Birthday', 'Phone', 'Addr', 'PolisN', 'IDPers'], axis=1,
                  inplace=True)
        xlsx.to_excel("t5.xlsx")
        dtb = pd.DataFrame()
        # недействующие
        print("7 этап из 9:")
        for ind in tqdm(range(len(xlsx))):
            id_excel = xlsx.iloc[ind]["6"]

            query = (f"""
                                 SELECT top 1 Pers.ENP,  Polis.PolisDateF,  _sCloseStatus.Name, Pers.IDPers
                   FROM Pers INNER JOIN
                   Polis ON Pers.IDPers = Polis.IDPers inner join
                   _sCloseStatus on _sCloseStatus.IDCloseStatus = Polis.IDCloseStatus
                   WHERE (Polis.PolisDateF IS not NULL) and Pers.Enp = '{id_excel}'
                   ORDER BY  Polis.PolisDateF Desc
                   """)
            dtb1 = pd.read_sql(query, eng)
            dtb = pd.concat([dtb, dtb1], ignore_index=True)
        xlsx = pd.concat([xlsx, dtb], axis=1)
        xlsx.to_excel("concat7.xlsx")

        print("8 этап из 9:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['6'] == xlsx.iloc[ind]['ENP'] and type(xlsx.iloc[ind1]['A51']) is not str and \
                        xlsx.iloc[ind1]['A63'] != 0:
                    xlsx.at[ind1, 'A57'] = xlsx.iloc[ind]['PolisDateF']
                    xlsx.at[ind1, 'A58'] = xlsx.iloc[ind]['Name']
                    xlsx.at[ind1, 'A59'] = xlsx.iloc[ind]['IDPers']
                    xlsx.at[ind1, "A63"] = 0
                    break
        xlsx.drop(columns=['ENP', 'PolisDateF', 'Name', 'IDPers'], axis=1, inplace=True)
        xlsx.to_excel("got5.xlsx")
        # выбор номеров
        c = []
        print("9 этап из 9:")
        for ind in tqdm(range(len(xlsx))):
            if type(xlsx.iloc[ind]['A58']) is str or xlsx.iloc[ind]['A63'] == 0:
                continue
            elif type(xlsx.iloc[ind]['10']) is not str and type(xlsx.iloc[ind]['A55']) is not str:
                xlsx.loc[ind, 'A60'] = None
            elif type(xlsx.iloc[ind]['10']) is str:
                b = ''.join(xlsx.iloc[ind]['10'].split())
                b = ''.join(b.split("+"))
                b = ''.join(b.split("("))
                b = ''.join(b.split(")"))
                b = ''.join(b.split("-"))
                match = re.findall(r'[78][9]\d{9}', b)
                rematch = re.findall(r'\d{6}', b)
                if match != c:
                    xlsx.at[ind, 'A60'] = "смс"
                    xlsx.at[ind, 'A61'] = match
                elif type(xlsx.iloc[ind]['A55']) is str:
                    a = ''.join(xlsx.iloc[ind]['A55'].split())
                    a = ''.join(a.split("+"))
                    a = ''.join(a.split("("))
                    a = ''.join(a.split(")"))
                    a = ''.join(a.split("-"))
                    a = ''.join(itertools.filterfalse(str.isalpha, a))
                    sot = re.findall(r'[9]\d{9}', a)
                    zvon = re.findall(r'd{6}?', a)
                    if sot != c:
                        xlsx.at[ind, 'A60'] = "смс"
                        xlsx.at[ind, 'A61'] = sot[0]
                    elif zvon != c:
                        xlsx.at[ind, 'A60'] = "звонок"
                        xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A55']
                    elif rematch != c:
                        xlsx.at[ind, 'A60'] = "звонок"
                        xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['10']

                    else:
                        xlsx.at[ind, 'A60'] = "письмо"
                else:
                    xlsx.at[ind, 'A60'] = "письмо"
            elif type(xlsx.iloc[ind]['A55']) is str:
                e = ''.join(xlsx.iloc[ind]['A55'].split())
                e = ''.join(e.split("+"))
                e = ''.join(e.split("("))
                e = ''.join(e.split(")"))
                e = ''.join(e.split("-"))
                e = ''.join(itertools.filterfalse(str.isalpha, e))
                sot = re.findall(r'[9]\d{9}', e)
                zvon = re.findall(r'\d{6}?', e)
                if sot != c:
                    xlsx.at[ind, 'A60'] = "смс"
                    xlsx.at[ind, 'A61'] = sot[0]
                elif zvon != c:
                    xlsx.at[ind, 'A60'] = "звонок"
                    xlsx.at[ind, 'A61'] = xlsx.iloc[ind]['A55']
                else:
                    xlsx.at[ind, 'A60'] = "письмо"
            else:
                xlsx.at[ind, 'A60'] = "письмо"
        for ind in tqdm(range(len(xlsx))):
            a = str(xlsx.iloc[ind]['A61'])
            a = ''.join(a.split())
            a = ''.join(a.split("["))
            a = ''.join(a.split("]"))
            a = ''.join(a.split("'"))
            a = ''.join(a.split("+"))
            a = ''.join(a.split("("))
            a = ''.join(a.split(")"))
            a = ''.join(a.split("-"))
            a = ''.join(itertools.filterfalse(str.isalpha, a))
            xlsx.at[ind, 'A61'] = a
        xlsx.rename(columns={
            'A51': "Фамилия ЗЛ",
            'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
            'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
            'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
            'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
            'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
            'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
            'A64': "Способ  информирования (по предыдущему информированию)",
            'A65': "Статус отправки",
            'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Вакцинация/Обработка/SMOVAC_informed.xlsx')
    if event == 'ivact':
        pathh = values["-FILE_PATH_INDEXVACT-"]
        xlsx = pd.read_excel(pathh, header=4, names=(
            "Index", "A2", "A3", "A4", "A5", "A90", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15",
            "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26",)
                             , dtype={"A90": str})
        dtb = pd.read_excel('Z:/Данные из ГИС ОМС/Вакцинация/Обработка/1.xlsx',
                            dtype={"6": str, "8": str, "A61": str})
        xlsx = xlsx[['Index', 'A90']]
        dtb.to_excel(dirname + "/Выходные файлы/SMO_changed1.xlsx")
        xlsx = pd.concat([xlsx, dtb], axis=1)
        print("1 этап из 1:")
        for ind1 in tqdm(range(len(xlsx))):
            for ind in range(len(xlsx)):
                if xlsx.iloc[ind1]['6'] == xlsx.iloc[ind]['A90']:
                    xlsx.at[ind1, '1'] = xlsx.iloc[ind]['Index']
                    break
        # xlsx.drop(columns=['To','Send At','Status', 'Seen At'],axis = 1, inplace=True)
        xlsx.drop(columns=['Index', 'A90'], axis=1, inplace=True)
        xlsx.rename(columns={'A1': "№ п/п", 'A6': "ЕНП", 'A8': "Номер бланка/полиса старого образца",
                             'A9': "Дата выдачи полиса",
                             'A10': "Группа", 'A11': "Телефон ЗЛ", 'A25': "Тип диспансеризации (2 половина 2021)",
                             'A30': "Наименование СП МО (адреса прикрепления)",
                             'A38': "Способ информирования_эл.почта", 'A39': "Способ информирования_телефон",
                             'A40': "Способ информирования_смс", 'A41': "Способ информирования_мессенджеры",
                             'A42': "Способ информирования_почта", 'A43': "Способ информирования_лично",
                             'A44': "Результат информирования", 'A45': "Действие СМО",
                             'A46': "Фамилия страхового представителя", 'A47': "Имя страхового представителя",
                             'A48': "Отчество страхового представителя", 'A49': "Телефон страхового представителя",
                             'A50': "Адрес страхового представителя",
                             'A51': "Фамилия ЗЛ",
                             'A52': "Имя ЗЛ", 'A53': "Отчество ЗЛ",
                             'A54': "Дата рождения ЗЛ", 'A55': "Телефонные номера ЗЛ",
                             'A56': "Адрес ЗЛ", 'A57': "Дата аннулирования полиса",
                             'A58': "Причина аннулирования", 'A59': "ID ЗЛ",
                             'A60': "Тип информирования", 'A61': "Телефон ЗЛ_осн",
                             'A62': "Дата последнего информирования в этом году", 'A63': "Вид информирования",
                             'A64': "Способ  информирования (по предыдущему информированию)",
                             'A65': "Статус отправки",
                             'A66': "Способ информирования (текущий)"}, inplace=True)
        xlsx.to_excel('Z:/Данные из ГИС ОМС/УД/Обработка/SMO_измененный_порядок.xlsx')
    if event in (None, 'Exit', 'Закрыть'):
        break