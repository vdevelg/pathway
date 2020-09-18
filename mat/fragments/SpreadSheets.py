import pyexcel
import os

def read(fileName):
    """
        Чтение из файла электронной таблицы в словарь
        { имяЛиста1:строка1[ ячейка1[], ... ], ... }
        и его заполнение
    """
    book_dict = pyexcel.get_book_dict(file_name=fileName)
    for listName in book_dict:
        # заполнение пустых ячеек в строке заголовка
        sheet_list = book_dict[listName]
        startTitle = ""
        for i in range(len(sheet_list[0])):
            title = sheet_list[0][i].strip()
            if title != "":
                startTitle = title
            sheet_list[0][i] = startTitle # замена текущего заголовка на обработанный методом strip()
            # заполнение ячеек в столбцах данных, содержащих ссылку "^" на расположенную выше ячейку
            startData = ""
            for j in range(1, len(sheet_list)):
                data = sheet_list[j][i].strip()
                if data != "^":
                    startData = data
                sheet_list[j][i] = startData
    return makePaths(book_dict)

def makePaths(bookDict):
    """ 
        Создание словаря Paths 
        { имяЛиста1:[ строкаЗаголовков[ текстоваяСтрока1, ... ], строкаПутей0[], ... ], ... }
        из словаря - представления книги электронных таблиц
        { имяЛиста1:строка1[ ячейка1[], ... ], ... }
    """
    Paths = {}
    for algorythm in bookDict:
        sheet_list = book_dict[algorythm]
        head = sheet_list[0]

        Paths[algorythm] = []
        pathList = Paths[algorythm]

        for i in range(1,len(sheet_list)):
            for j in range(len(head)):
                curPart = sheet_list[i][j]
                if head[j].endswith("Dir"):
                    startHead = 
                elif head[j].endswith("File"):
                    
                elif head[j].endswith("Type"):






        for head in bookDict[algorythm][0]:
            if head not in Paths[algorythm]:
                Paths[algorythm][head] = ""

        for datastring in bookDict[algorythm][1:]:
            head = bookDict[algorythm][0][0].upper()
            path = ""
            for i in range( len(datastring) ):
                curHead = bookDict[algorythm][0][i].upper()
                if curHead != head:
                    Paths[algorythm][head] = path
                    path = ""
                    head = curHead
                path = os.path.join( path, datastring[i] )
    return Paths

def Хлам():
    def a():
    PathsDict = SpreadSheets.read(values["#gFile"])
    print(PathsDict)
    dataOfTables = {}
    for algorythm in PathsDict:
        dataOfTables[algorythm] = []
        ttable = list( PathsDict[algorythm].values() )
        print(ttable)
        for i in range( len( ttable[0] ) ):
            stringOfTable = []
            for j in range( len( ttable ) ):
                stringOfTable.append( ttable[j][i] )
            dataOfTables[algorythm].append( stringOfTable )

# paths = read("D:/Users/VG/proj/p_Python/paths/материалы/paths.ods")
# for i in paths:
#     for j in paths[i]:
#         print(j)


# a = ["123","234","345","456","567","678"]
# d = ""
# for i in a:
#     d = os.path.join(d,i)

# print(d)
