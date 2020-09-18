import comtypes.client
import os.path
import re
import pyexcel

"""
    Способ хранения и извлечения переменных в документах Word:
    https://support.microsoft.com/ru-ru/help/306281/how-to-store-and-retrieve-variables-in-word-documents
"""

def counter(start=0, step=1, number=1):
    """ Возвращает функцию-счётчик, которая, в свою очередь, 
        возвращает значение счётчика, начиная со start, с шагом step, 
        приращение значения происходит при количестве number вызовов функции-счётчика
    """
    i = 0 # Количество вызовов с последнего сброса (но это не точно)
    count = start # Переменная счётчика
    def incrementer():
        nonlocal i, count, step, number
        i += 1
        if i > number:
            i = 1
            count += step
        return count
    return incrementer

def keygen(path):
    """
        Функция на основе одного абсолютного пути к файлу, в котором нужно изменить значение переменной, 
        должна возвратить уникальный для данной группы файлов ключ.
        К сожалению, нет общего алгоритма для реализации данной функции, 
        поэтому для каждого проекта нужно править код этой функции.
    """
    global call_num
    algorithm = 0 # можно выбрать алгоритм из уже реализованных

    if   algorithm == 0:
        key = call_num()
    elif algorithm == 1:
        dir_name = os.path.split(os.path.dirname(path))[1] # имя директории в которой содержится файл назначения
        result = re.findall(r'\d+', dir_name) # список всех чисел в имени директории
        key = result[0] # выбор первого числа в качестве ключа
    elif algorithm == 2:
        file_name = os.path.split(path)[1] # полное имя файла с расширением имени
        short_file_name = file_name.split(".")[0] # разделение имени и расширения имени, выбор первой части
        key = short_file_name.split(" ")[0] # выбор части строки до первого пробела
    else:
        print("Неверно указан алгоритм для функции keygen, либо его не существует.")
    key = str(key)
    return key

def readDocVars(path):
    """ 
        Печатает и возвращает список переменных документа
    """
    global Word
    doc = Word.Documents.Open(path) # открытие документа
    print("Документ содержит переменные:")
    DocVariables = {}
    for var in doc.Variables:
        DocVariables[var.Name] = var.Value
        print(var.Name, "=", var.Value)
    doc.Close(0) # закрытие документа
    return DocVariables

def setDocVars(path, DocVariables):
    """ 
        Установка переменных документа
    """
    global Word
    doc = Word.Documents.Open(path) # открытие документа
    print("Запись переменных:")
    for VariableName in DocVariables:
        print(VariableName, "=", DocVariables[VariableName])
        doc.Variables.Item(VariableName).Value = DocVariables[VariableName]
    """
        Для обновления всех полей в документе (кроме поля содержания) 
        используется режим предварительного просмотра, при переходе в который
        автоматически обновляются все поля (это необходимо настроить в Word)
    """
    doc.ScreenUpdating = False
    doc.PrintPreview()
    doc.ClosePrintPreview()
    doc.ScreenUpdating = True
    """
        Содержание можно принудительно обновить отдельной командой.
        В документе может быть несколько содержаний, 
        поэтому в скобках нужно указать порядковый номер содержания. 
        Если содержание одно, то всё-равно в скобках надо указать порядковый номер:
    """
    # doc.TablesOfContents(1).Update()
    doc.Fields.Update() # Обновление всех полей в основном тексте документа
    doc.Save()
    doc.Close(0) # закрытие документа

def delAllDocVars(path):
    """
        Удаление всех переменных документа
    """
    global Word
    doc = Word.Documents.Open(path) # открытие документа
    for var in doc.Variables:
        doc.Variables(var.Name).Delete()
    doc.Save()
    doc.Close(0) # закрытие документа

def read_docVars_data(file_name):
    var_book = pyexcel.get_book(file_name=file_name)
    keys_name = ""
    DocumentVariables = {}
    for sheet in var_book:
        if sheet.name.startswith("-"):
            continue
        for column in sheet.columns():
            if str(column[0]).strip() in ("", keys_name):
                continue
            for i in range(len(column)):
                cur_val = str(column[i]).strip()
                if cur_val != "^":
                    tit_val = cur_val
                column[i] = tit_val
            DocumentVariables[column[0]] = column[1:]
        if keys_name == "":
            keys_name = var_book[sheet.name][0, 0]
    return DocumentVariables, keys_name

def docVars(ProcData):
    """
        Основная вызываемая функция
    """
    global Word, call_num

    call_num = counter(start=1)
    Word = comtypes.client.CreateObject("Word.Application") # открытие приложения Word
    # Word.Visible = True # Отображение окна приложения

    prev_in_path = ""
    prev_key = ""
    for row in ProcData:
        cur_in_path = row[0]
        if cur_in_path != prev_in_path:
            # Чтение входного файла и формирование словаря:
            DocumentVariables, keys_name = read_docVars_data(cur_in_path)
        prev_in_path = cur_in_path

        cur_out_path = row[1]
        # Получение ключа и формирование словаря DocVariables:
        cur_key = keygen(cur_out_path)
        if cur_key != prev_key:
            for row_num in range(len(DocumentVariables[keys_name])):
                if DocumentVariables[keys_name][row_num] == cur_key:
                    key_num = row_num
            # Формируем словарь DocVariables:
            DocVariables = {}
            for var_name in DocumentVariables:
                DocVariables[var_name] = DocumentVariables[var_name][key_num]
        prev_key = cur_key
        # Непосредственно обработка файла по алгоритму:
        # readDocVars(cur_out_path)
        setDocVars(cur_out_path, DocVariables)
        # delAllDocVars(cur_out_path)

    Word.Quit() # закрытие приложения

def docVars_print(ProcData):
    """
        Тестовая функция
    """
    print()
    print("Привет, меня зовут docVars")
    print("Мне передали аргументы:")
    for i in ProcData:
        print(i)
    print()
