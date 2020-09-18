import sys
import os
import pyexcel as pycel
from pathtools import splitPDF, xlsx2pdf

xlsx2pdf.proc()
splitPDF.proc()

def DSF(FLNM):
    book_dict = pycel.get_book_dict(file_name=FLNM)
    sheet_list = book_dict["xlsx2pdf"]
    
    starttitle = ""
    for i in range(len(sheet_list[0])):
        title = sheet_list[0][i].strip()
        if title != "":
            starttitle = title
        sheet_list[0][i] = starttitle
        startdata = ""
        for j in range(1, len(sheet_list)):
            data = sheet_list[j][i].strip()
            if data != "":
                startdata = data
            sheet_list[j][i] = startdata
    return sheet_list

def proc_loop(sheet_list):
    while True:
        execute()
        



if len(sys.argv) < 1:
    print("Fatal Error! Ahtung!")
if len(sys.argv) == 1:
    print("Укажите место сохранения файла-примера")
if len(sys.argv) == 2:
    WD = os.path.dirname(sys.argv[1]) # директория файла электронной таблицы - рабочая директория
    #os.chdir(WD) # установка рабочей директории в качестве текущей
    DSF(sys.argv[1])
if len(sys.argv) > 2:
    print("Fatal Error! Ahtung!")



input("\n\n\nНажмите Enter для выхода")
