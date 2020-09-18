# Программа конвертирует файлы ".docx" в файлы ".pdf"

import os
import comtypes.client
import sys

global Word, doc

def docx_to_pdf(in_file_path, out_file_path):
    """
        Конвертирует файл формата .docx расположенный по пути in_file_path
        в файл формата .pdf расположенный по пути out_file_path
    """
    doc = Word.Documents.Open(in_file_path) # откркрытие документа
    doc.SaveAs(out_file_path, FileFormat=17) # сохранение в формате .pdf
    doc.Close(0) # закрытие документа

##____________________________________end_defs____________________________________


for i in range(len(sys.argv)):
    if i == 0:
        script_name = os.path.split(sys.argv[i])[1]
    elif i == 1:
        os.chdir(sys.argv[i])

start_text = script_name.replace(" docx to pdf.py", "")
end_text = ".docx"

Word = comtypes.client.CreateObject("Word.Application") # открытие приложеня Word
##Word.Visible = True

for name in os.listdir(path="."):
    if name.startswith(start_text) and name.endswith(end_text):
        in_file_path = os.path.join(os.getcwd(), name)
        out_file_path = os.path.join(os.getcwd(), name.replace(".docx", ".pdf"))
        docx_to_pdf(in_file_path, out_file_path)
        print(". ")
        
Word.Quit() # закрытие приложеня Word

##print("Программа завершила работу")
##input("Нажмите Enter для выхода")
