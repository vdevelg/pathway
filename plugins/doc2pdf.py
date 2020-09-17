import comtypes.client

def doc_to_pdf(in_file_path, out_file_path):
    global Word
    doc = Word.Documents.Open(in_file_path) # открытие документа
    doc.SaveAs(out_file_path, FileFormat=17) # сохранение в формате .pdf
    doc.Close(0) # закрытие документа

def doc2pdf(ProcData):
    global Word
    Word = comtypes.client.CreateObject("Word.Application") # открытие приложеня Word
    # Word.Visible = True # отображение окна приложения
    for row in ProcData:
        doc_to_pdf(row[0], row[1])
    Word.Quit() # закрытие приложения

def doc2pdf_print(ProcData):
    print()
    print("Привет, меня зовут doc2pdf")
    print("Мне передали аргументы:")
    for i in ProcData:
        print(i)
    print()
