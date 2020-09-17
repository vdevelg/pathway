import os

def preproc(sheet_list, CWD):
    """ 
        Предобработка таблицы данных
    """
    # Чёткое очерчивание границ таблицы:
    col_num = 0
    while col_num < len( sheet_list[0] ):
        if sheet_list[0][col_num].strip() == "":
            for row_num in range( len( sheet_list ) ):
                del sheet_list[row_num][col_num]
        else:
            col_num += 1
    row_num = 0
    while row_num < len( sheet_list ):
        if sheet_list[row_num][0].strip() == "":
            del sheet_list[row_num]
        else:
            row_num += 1
    # Заполнение ячеек в строке заголовков, содержащих ссылку "<" на расположенную слева ячейку: 
    startTitle = ""
    for col_num in range( len( sheet_list[0] ) ):
        title = sheet_list[0][col_num].strip()
        if title != "<":
            startTitle = title
        sheet_list[0][col_num] = startTitle # замена текущего заголовка на обработанный методом strip()
        # Заполнение ячеек в столбцах данных, содержащих ссылку "^" на расположенную выше ячейку:
        startData = ""
        for row_num in range(1, len(sheet_list)):
            data = str(sheet_list[row_num][col_num]).strip().replace("<CWD>", CWD) # обработка данных в ячейке
            if data != "^":
                startData = data
            sheet_list[row_num][col_num] = startData # замена текущих данных на обработанные

def append_file_paths(dir_path, file_names, file_types, file_paths):
    lfn = len(file_names)
    lft = len(file_types)
    if   lfn == 1 and lft == 1:
        file_paths.append( os.path.join( dir_path, file_names[0] + file_types[0] ) )
    elif lfn >  1 and lft == 1:
        for i in range(lfn):
            file_paths.append( os.path.join( dir_path, file_names[i] + file_types[0] ) )
    elif lfn == 1 and lft >  1:
        for i in range(lft):
            file_paths.append( os.path.join( dir_path, file_names[0] + file_types[i] ) )
    elif lfn == lft and lfn != 0:
        for i in range(lfn):
            file_paths.append( os.path.join( dir_path, file_names[i] + file_types[i] ) )
    else:
        print("!Ошибка формата входного файла!")

def make_paths(sheet_list):
    """
        Тихий ужас... Функция парсит формат файла с путями.
    """
    PKM = { "InMarker" : "In", 
            "OutMarker" : "Out" } # Part kind markers
    PTM = { "DirMarker" : "Dir", 
            "NameMarker" : "Fname", 
            "TypeMarker" : "Ftype" } # Part type markers

    paths = []
    for row_num in range(1, len(sheet_list)): # Перебор строк в таблице
        prevPartKind = "" # Вид предыдущей части (ячейки таблицы)
        prevPartType = "" # Тип предыдущей части (ячейки таблицы)
        dir_path = "" # Concatenated DirMarkered strings
        file_names = [] # Appended NameMarkered strings
        file_types = [] # Appended TypeMarkered strings
        file_paths = [] # Appended DirMarkered strings of current kind: InMarker or OutMarker
        for col_num in range(len(sheet_list[0])): # Перебор столбцов в таблице
            curPart = sheet_list[row_num][col_num] # Текущая часть
            title = sheet_list[0][col_num] # Заголовок текущей части
            # Определение вида и типа текущей части:
            for marker in PKM:
                if title.startswith(PKM[marker]):
                    curPartKind = PKM[marker]
            for marker in PTM:
                if title.endswith(PTM[marker]):
                    curPartType = PTM[marker]
            # Обработка части в строке согласно её типу:
            if   (prevPartType, curPartType) == ("", PTM["DirMarker"]):
                dir_path = os.path.join( dir_path, curPart )
            elif (prevPartType, curPartType) == (PTM["DirMarker"], PTM["DirMarker"]):
                dir_path = os.path.join( dir_path, curPart )
            elif (prevPartType, curPartType) == (PTM["DirMarker"], PTM["NameMarker"]):
                file_names = [curPart]
            elif (prevPartType, curPartType) == (PTM["DirMarker"], PTM["TypeMarker"]):
                print("! Ошибка формата входного файла. №1 !")
            elif (prevPartType, curPartType) == (PTM["NameMarker"], PTM["DirMarker"]):
                print("! Ошибка формата входного файла. №2 !")
            elif (prevPartType, curPartType) == (PTM["NameMarker"], PTM["NameMarker"]):
                file_names.append(curPart)
            elif (prevPartType, curPartType) == (PTM["NameMarker"], PTM["TypeMarker"]):
                file_types = [curPart]
            elif (prevPartType, curPartType) == (PTM["TypeMarker"], PTM["DirMarker"]):
                append_file_paths(dir_path, file_names, file_types, file_paths)
                file_names = []
                file_types = []
                dir_path = curPart
            elif (prevPartType, curPartType) == (PTM["TypeMarker"], PTM["NameMarker"]):
                append_file_paths(dir_path, file_names, file_types, file_paths)
                file_names = [curPart]
                file_types = []
            elif (prevPartType, curPartType) == (PTM["TypeMarker"], PTM["TypeMarker"]):
                file_types.append(curPart)
            else:
                print("! Неизвестный формат входного файла !", prevPartType, curPart)

            # Обработка части в строке согласно её виду:
            if   (prevPartKind, curPartKind) == ("", PKM["InMarker"]):
                pass
            elif (prevPartKind, curPartKind) == (PKM["InMarker"], PKM["OutMarker"]):
                if prevPartType == PTM["TypeMarker"]:
                    in_paths = file_paths # Формирование списка путей входных файлов
                    file_paths = []
                else:
                    print("! Ошибка формата входного файла. №3 !")
            elif (prevPartKind, curPartKind) == (PKM["OutMarker"], PKM["InMarker"]):
                print("! Ошибка формата входного файла. №4 !")

            prevPartKind = curPartKind
            prevPartType = curPartType

        if curPartKind != PKM["OutMarker"] and curPartType != PTM["TypeMarker"]:
            print("! Ошибка формата входного файла. №5 !")
        else:
            append_file_paths(dir_path, file_names, file_types, file_paths)
            out_paths = file_paths  # Формирование списка путей выходных файлов

        paths.append( [in_paths, out_paths] )
    return paths

def tabulate(file_paths):
    """ 
        Приводит формат таблицы с произвольным количеством столбцов:
        file_paths = [ [ [ "in_path_1", "in_path_2", ... ], [ "out_path_1", ... ] ],
                       [ [ "in_path_1", "in_path_2", ... ], [ "out_path_1", ... ] ] ... ]
        к формату таблицы с двумя столбцами:
        TabData = [ [ "in_path_1", "out_path_1"],
                    [ "in_path_2", "out_path_2"], ... ]
    """
    # Добавление элементов, содержащих пустую строку, к списку одного из видов 
    # для выравнивания длины списков всех видов:
    for group in file_paths: # Перебор по строкам (группам)
        lens = [] # Длины списков путей различных видов
        for popul in group: # Перебор путей одного вида: InMarker или OutMarker
            lens.append( len(popul) )
        max_len = max(lens)
        for col_num in range( len(group) ):
            group[col_num] += ["" for i in range( max_len - lens[col_num] ) ]
    # Формирование представления в несколько столбцов - по количеству видов путей к файлам(входные или выходные):
    TabData = [] # список(таблица) нового представления
    group_nums = [] # номера групп каждой строки таблицы
    for group_num in range( len(file_paths) ):
        for item_num in range( len( file_paths[group_num][0] ) ):
            row = []
            for popul in file_paths[group_num]:
                row.append(popul[item_num])
            TabData.append(row)
            group_nums.append(group_num)
    return TabData, group_nums

def tabdisp(TabData, group_nums):
    """ 
        Использует формат таблицы TabData, используемый для удобной обработки данных,
        для формирования формата таблицы TabDisp (для удобства восприятия) и передачи в графический интерфейс
    """
    # Добавление информации о группах:
    row_colors = []
    color = "#9FB8AD" # цвет нечётных групп
    TabDisp = []
    prev_group_num = -1
    for row_num in range( len( group_nums ) ):
        # Формирование списка кортежей для раскрашивания таблицы по группам:
        if group_nums[row_num] % 2 == 0:
            row_colors.append( ( row_num, color ) )
        # Добавление в таблицу столбца с номерами групп:
        if prev_group_num != group_nums[row_num]:
            TabDisp.append( [ group_nums[row_num]+2 ] + TabData[row_num] )
            prev_group_num = group_nums[row_num]
        else:
            TabDisp.append( [""] + TabData[row_num] )
    return TabDisp, row_colors

def process(Data, algorithm):
    """
        Интерфейс всех плагинов общий. 
        Утилита пакетной обработки "Обработчик путей" (данная программа) 
        передаёт в плагин список списков Data. 
        Каждый из вложенных списков содержит два элемента: входной и выходной пути.
    """
    if   algorithm == "copyFiles":
        from plugins.copyFiles import copyFiles
        copyFiles(Data)
    elif algorithm == "docVars":
        from plugins.docVars import docVars
        docVars(Data)
    elif algorithm == "doc2pdf":
        from plugins.doc2pdf import doc2pdf
        doc2pdf(Data)
    elif algorithm == "dwgVars":
        from plugins.dwgVars import dwgVars
        dwgVars(Data)
    elif algorithm == "dwg2pdf":
        from plugins.dwg2pdf import dwg2pdf
        dwg2pdf(Data)
    elif algorithm == "joinPDF":
        from plugins.joinPDF import joinPDF
        joinPDF(Data)
    elif algorithm == "setMarg":
        from plugins.setMarg import setMarg
        setMarg(Data)
    else:
        return 1
