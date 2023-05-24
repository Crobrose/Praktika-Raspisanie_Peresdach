import gspread
import json
import datetime

from gspread import  Client, Spreadsheet, Worksheet
from os import read
from operator import itemgetter, attrgetter

MasElemResult = [] #1 Половина
MasElemResult2 = [] #2 Половина
year = datetime.datetime.now().year - 2
month = datetime.datetime.now().month
period = 1 if(month > 9) else 2 # Определение какой сейчас семак
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1aQFuvdPUdLepg6LExVAdefpuM1VHRSHuSCT6bp_E6MI/edit#gid=0"

def sort_json_file():
    myfile = "Институт информатики и вычислительной техники_" + str(year) + " - " + str(year + 1) + "_Дисц.json"
    if period == 2:
        myfile2 = "Институт информатики и вычислительной техники_" + str(year + 1) + " - " + str(year + 2) + "_Дисц.json"

    with open(myfile, "r", encoding="utf-8-sig") as Filem:
        MasElemFirstYear = json.load(Filem)
    if period == 2:
        with open(myfile2, "r", encoding="utf-8-sig") as Filem:
            MasElemSecondYear = json.load(Filem)

    for i in range(1,3):
        kol_grup_in_str = 1

        for elem in MasElemFirstYear:
            for elems in MasElemFirstYear:
                if elem["Семестр"] == elems["Семестр"] and elem["Наименование"] == elems["Наименование"] and elem[
                    "Преподаватели"] == elems["Преподаватели"] and elem["ВидКонтроля"] == elems["ВидКонтроля"]:
                    if elem["Группы"] != elems["Группы"]:
                        elem["Группы"] += ", "
                        if kol_grup_in_str == 2:
                            kol_grup_in_str = 0
                            elem["Группы"] += "\n"
                        elem["Группы"] += elems["Группы"]
                        kol_grup_in_str += 1

                        MasElemFirstYear.remove(elems)

        kol_grup_in_str = 1

        for elem in MasElemSecondYear:
            for elems in MasElemSecondYear:
                if elem["Семестр"] == elems["Семестр"] and elem["Наименование"] == elems["Наименование"] and elem[
                    "Преподаватели"] == elems["Преподаватели"] and elem["ВидКонтроля"] == elems["ВидКонтроля"]:
                    if elem["Группы"] != elems["Группы"]:
                        elem["Группы"] += ", "
                        if kol_grup_in_str == 2:
                            kol_grup_in_str = 0
                            elem["Группы"] += "\n"
                        elem["Группы"] += elems["Группы"]
                        kol_grup_in_str += 1
                        MasElemSecondYear.remove(elems)

    for elem in MasElemFirstYear:
        if elem['Семестр'] == "Первый семестр":
            elem['Семестр'] = "1"
        elif elem['Семестр'] == "Второй семестр":
            elem['Семестр'] = "2"
        elif elem['Семестр'] == "Третий семестр":
            elem['Семестр'] = "3"
        elif elem['Семестр'] == "Четвертый семестр":
            elem['Семестр'] = "4"
        elif elem['Семестр'] == "Пятый семестр":
            elem['Семестр'] = "5"
        elif elem['Семестр'] == "Шестой семестр":
            elem['Семестр'] = "6"
        elif elem['Семестр'] == "Седьмой семестр":
            elem['Семестр'] = "7"
        elif elem['Семестр'] == "Восьмой семестр":
            elem['Семестр'] = "8"
        elif elem['Семестр'] == "Девятый семестр":
            elem['Семестр'] = "9"
        elif elem['Семестр'] == "Десятый семестр":
            elem['Семестр'] = "10"
        elif elem['Семестр'] == "Одиннадцатый семестр":
            elem['Семестр'] = "11"
    for elem in MasElemSecondYear:
        if elem['Семестр'] == "Первый семестр":
            elem['Семестр'] = "1"
        elif elem['Семестр'] == "Второй семестр":
            elem['Семестр'] = "2"
        elif elem['Семестр'] == "Третий семестр":
            elem['Семестр'] = "3"
        elif elem['Семестр'] == "Четвертый семестр":
            elem['Семестр'] = "4"
        elif elem['Семестр'] == "Пятый семестр":
            elem['Семестр'] = "5"
        elif elem['Семестр'] == "Шестой семестр":
            elem['Семестр'] = "6"
        elif elem['Семестр'] == "Седьмой семестр":
            elem['Семестр'] = "7"
        elif elem['Семестр'] == "Восьмой семестр":
            elem['Семестр'] = "8"
        elif elem['Семестр'] == "Девятый семестр":
            elem['Семестр'] = "9"
        elif elem['Семестр'] == "Десятый семестр":
            elem['Семестр'] = "10"
        elif elem['Семестр'] == "Одиннадцатый семестр":
            elem['Семестр'] = "11"
    #for kus in MasElemFirstYear:
        #print(kus['Семестр'],kus['Наименование'],kus['Преподаватели'],kus['Группы'])

    for elem in MasElemFirstYear:
        if elem["Семестр"] in {"1", "3", "5", "7", "9", "11"}:
            if period == 1:
                MasElemResult.append(elem)
        else:
            if period == 1:
                MasElemResult2.append(elem)
            else:
                MasElemResult.append(elem)


    if period == 2:
        for elem in MasElemSecondYear:
            if elem['Семестр'] in {"1", "3", "5", "7", "9", "11"}:
                MasElemResult2.append(elem)

    lookingfor = "("
    for elem in MasElemResult:
        kaf = elem['Кафедра']
        if(kaf == "Кафедра иностранных и русского языков"):
            elem['Кафедра'] = "ИиРЯ"
        if(kaf == "Кафедра физики"):
            elem['Кафедра'] = "Физики"
        for c in range(0, len(kaf)):
            if kaf[c] == lookingfor:
                elem['Кафедра'] = kaf[c+1:len(kaf)-1]

    for elem in MasElemResult2:
        kaf = elem['Кафедра']
        if (kaf == "Кафедра иностранных и русского языков"):
            elem['Кафедра'] = "ИиРЯ"
        if (kaf == "Кафедра физики"):
            elem['Кафедра'] = "Физики"
        for c in range(0, len(kaf)):
            if kaf[c] == lookingfor:
                elem['Кафедра'] = kaf[c+1:len(kaf)-1]

    MasElemResult.sort(key=itemgetter('Семестр'))
    MasElemResult2.sort(key=itemgetter('Семестр'))
def work_with_google_sheets():
    header_row = ["Наименование", "ВидКонтроля", "Кафедра", "Преподаватели", "Группы", "Дата", "Время", "ауд/ссылка на платформу", "Доп. информация", "Семестр"]
    rows1 = [header_row]
    rows1Mag = [header_row]
    rows2 = [header_row]
    rows2Mag = [header_row]
    txt = ""

    MasElemResultNew = []
    MasElemResultNew2 = []

    for elem in MasElemResult:
        if elem['Семестр'] in {"2", "4", "6"}:
            rows1.append([
                txt.join(elem.get(key, ""))
                for key in header_row
            ])

        if elem['Семестр'] in {"10"}:
            rows1Mag.append([
                txt.join(elem.get(key, ""))
                for key in header_row
            ])
    rows1.append([])
    for elem in MasElemResult2:
        if elem['Семестр'] in {"1", "3", "5", "7"}:
            rows2.append([
                txt.join(elem.get(key, ""))
                for key in header_row
            ])

        if elem['Семестр'] in {"9", "11"}:
            rows2Mag.append([
                txt.join(elem.get(key, ""))
                for key in header_row
            ])
    rows2.append([])
    #for kus in rows:
     #   print(kus)


    gc: Client = gspread.service_account("./service_account.json")
    sh: Spreadsheet = gc.open_by_url(SPREADSHEET_URL)

    worksheet_list = sh.worksheets()
    worksheet_index = 0

    if period == 2:
        per1 = " летней "
        per2 = " зимней "
    else:
        per1 = " зимней "
        per2 = " летней "

    for works in worksheet_list:
        worksheet_index = worksheet_index + 1

    if(worksheet_index == 1):
        sh.add_worksheet("Тест", rows=1, cols=1)

    if(worksheet_index > 2):
        for i in range(1, worksheet_index-1):
            sh.del_worksheet(sh.sheet1)

    sh.del_worksheet(sh.sheet1)
    comments_ws = sh.add_worksheet("Пересдачи" + per2 + "сессии " + str(year + 1) + " - " + str(year + 2), rows=1, cols=len(header_row))
    comments_ws.insert_rows(rows2Mag)
    comments_ws.insert_rows([])
    comments_ws.insert_rows(rows2)


    sh.del_worksheet(sh.sheet1)
    comments_ws = sh.add_worksheet("Пересдачи" + per1 + "сессии " + str(year) + " - " + str(year + 1), rows=1, cols=len(header_row))
    comments_ws.insert_rows(rows1Mag)
    comments_ws.insert_rows([])
    comments_ws.insert_rows(rows1)


    cell_format = {
        "backgroundColor": {
            "red": 1,
            "green": 0.8,
            "blue": 0.2
        },
        "textFormat": {
            "fontSize": 13,
            "bold": True
        }
    }
    worksheet1 = sh.get_worksheet(0)
    worksheet2 = sh.get_worksheet(1)
    worksheet1.format("A1:J1", cell_format)
    worksheet2.format("A1:J1", cell_format)
def main():
    sort_json_file()
    work_with_google_sheets()

main()