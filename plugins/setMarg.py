import comtypes.client

def setMargins(out_file_path, margins):
    global Word
    doc = Word.Documents.Open(out_file_path) # открытие документа
    doc.PageSetup.TopMargin = Word.MillimetersToPoints(margins[0][0])
    doc.PageSetup.BottomMargin = Word.MillimetersToPoints(margins[0][1])
    doc.PageSetup.LeftMargin = Word.MillimetersToPoints(margins[1][0])
    doc.PageSetup.RightMargin = Word.MillimetersToPoints(margins[1][1])
    doc.Save()
    doc.Close(0) # закрытие документа

def setMarg(ProcData):
    global Word
    margins = ((30, 20), (25, 10))
    Word = comtypes.client.CreateObject("Word.Application") # открытие приложеня Word
    # Word.Visible = True # отображение окна приложения
    for row in ProcData:
        setMargins(row[1], margins)
    Word.Quit() # закрытие приложения

def setMarg_print(ProcData):
    print()
    print("Привет, меня зовут setMarg")
    print("Мне передали аргументы:")
    for i in ProcData:
        print(i)
    print()
