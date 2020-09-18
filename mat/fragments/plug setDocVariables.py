import comtypes.client

def keygen():
    

"""
    Способ хранения и извлечения переменных в документах Word:
    https://support.microsoft.com/ru-ru/help/306281/how-to-store-and-retrieve-variables-in-word-documents
"""


in_file_path = r"D:\dev\Pathway\XXXX-XXXX-XXXXX.ТЧ (XXX-XX Текстовая часть).docx" # путь к файлу документа
Word = comtypes.client.CreateObject("Word.Application") # открытие приложения Word
Word.Visible = True # отображение окна приложения Word

doc = Word.Documents.Open(in_file_path) # открытие документа

DocVariables = {"TB_regul" : "Пупкин", "TB_regul_data" : "02.03.2020"}
# Задание переменной с именем Field значения "Field value":
for VariableName in DocVariables:
    doc.Variables.Item(VariableName).Value = DocVariables[VariableName]

# Перебор всех переменных и распечатка их имён и значений:
for var in doc.Variables:
    print("Имя: ", var.Name, "Значение: ", var.Value, sep="")

"""
    Для обновления всех полей в документе (кроме поля содержания) 
    используется режим предварительного просмотра, при переходе в который
    автоматически обновляются все поля
"""
doc.ScreenUpdating = False # Отключение обновления экрана
doc.PrintPreview() # Предварительный просмотр
doc.ClosePrintPreview() # Закрыть предварительный просмотр
doc.ScreenUpdating = True # Обновить экран

"""
    Содержание можно принудительно обновить следующей командой.
    В документе может быть несколько содержаний, 
    поэтому в скобках нужно указать порядковый номер содержания. 
    Если содержание одно, то всё-равно в скобках надо указать порядковый номер:
"""
# doc.TablesOfContents(1).Update()

# doc.Fields.Update() # Обновление всех полей в основном тексте документа

if doc.Saved == False: doc.Save()
# doc.SaveAs2(FileName="newname.pdf", FileFormat=17) # FileFormat=17 - PDF файл (переменная wdFormatPDF в VBA)

# doc.Close(0) # закрытие документа
# Word.Quit() # закрытие приложения Word
