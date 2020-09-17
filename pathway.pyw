import PySimpleGUI as sg
import pyexcel
import os
import lang.lang as lang
from lang.lang import GT, rep
from pathway_handlers import preproc, make_paths, tabulate, tabdisp, process



__progName__ = "Обработчик путей" # Name of programm
__version__ = "2020_03_04" # Version of programm
sg.theme('System Default 1') # Theme: "Default 1", "DarkTanBlue", "System Default 1"...

LANGUAGE = "RUS" # Code of language: "RUS", "ENG", "SWA"...
lang.read_lang_file(LANGUAGE)

algorithms = [] # Described algorithms
plugins = ["copyFiles", "docVars", "doc2pdf", "dwgVars", "dwg2pdf", "joinPDF", "setMarg"] # available plugins

# TableContent = [ [GT("Header1"), GT("Header2")] ] + \
#                [ ["Входной файл {}".format(nom_file), "Выходной файл {}".format(nom_file)] for nom_file in range(55) ]
# TabDisplay = [ [nom_file+2, "Входной файл {}".format(nom_file), "Выходной файл {}".format(nom_file)] for nom_file in range(55) ]

TabHeader = ["#", GT("Header1"), GT("Header2")]

layout = [
    [
        sg.Text( GT("File"), size=( max( len(GT("File") ), len( GT("Algorithm") ) ), 1 ) ), 
        sg.InputText("", key="#gFile", size=(140, 1)), 
        sg.FileBrowse(GT("Open"), key="#gBFile", size=(len(GT("Open"))+3, 1), target="#gBFile", enable_events=True)
    ], 
    [
        sg.Text( GT("Algorithm"), size=( max( len(GT("File") ), len( GT("Algorithm") ) ), 1 ) ), 
        sg.DropDown(algorithms, size=(20, 1), key="#gAlgorithms", enable_events=True)
    ], 
    [
        sg.Table( "", headings=TabHeader, 
                  key="#gTable", 
                  col_widths=[3, 62, 62], 
                  auto_size_columns=False, num_rows=20, 
                  justification="left", 
                  vertical_scroll_only=False, 
                  pad=((0,0),0)
                )
    ], 
    [
        # sg.Image( r".\logo\vdevel.png", size=(55, 20) ), 
        sg.Button( GT("SelectGroups"), size=(int(1.15*max(len(GT("SelectGroups")), len(GT("ProcSelected")))), 2), pad=((0,33),(20,10))), 
        sg.Button( GT("ProcSelected"), size=(int(1.15*max(len(GT("SelectGroups")), len(GT("ProcSelected")))), 2), pad=((0,33),(20,10))), 
        sg.Button( GT("ProcAll"), size=(int(1.15*len(GT("ProcAll"))), 1), pad=((100,380), (20,10))), 
        sg.Button( GT("Exit"), size=(int(2*len(GT("Exit"))), 2), pad=((0,0), (20,10)))
    ]
]



# -----------------------------------------------------------------------------------
window = sg.Window("{0}:  вер. {1}".format(__progName__, __version__), layout)
# -----------------------------------------------------------------------------------



while True:
    event, values = window.read()
    print(event, values)

    if   event in ( None, GT("Exit") ):
        break

    elif event == "#gBFile":
        if values["#gBFile"] == "":
            source_file = os.path.normpath(values["#gFile"]) if os.path.normpath(values["#gFile"]) != "." else ""
        else:
            source_file = os.path.normpath(values["#gBFile"])
        window["#gFile"].Update(source_file)
        if os.path.isfile(source_file):
            book_dict = pyexcel.get_book_dict(file_name=source_file)
            algorithms = []
            for algorythm in list(book_dict.keys()):
                if algorythm.startswith("-"):
                    continue
                algorithms.append(algorythm)
            window["#gAlgorithms"].Update(values=algorithms)
        else:
            sg.Popup(GT("FileNotExist"), title=GT("Error"), keep_on_top=True)

    elif event == "#gAlgorithms":
        if values["#gAlgorithms"] in plugins:
            sheet_list = book_dict[ values["#gAlgorithms"] ]
            CWD = os.path.dirname( values["#gFile"] )
            preproc(sheet_list, CWD)
            paths = make_paths(sheet_list)
            TabData, group_nums = tabulate(paths)
            TabDisplay, row_colors = tabdisp(TabData, group_nums)
            window["#gTable"].Update( TabDisplay, row_colors=row_colors )
        else:
            sg.Popup( GT("PluginNotFound").format(values["#gAlgorithms"]), title=GT("Error"), keep_on_top=True)

    elif event == GT("SelectGroups"):
        if values["#gTable"] == []:
            sg.Popup(GT("SelectTableRow"), title=GT("Error"), keep_on_top=True)
        else:
            # Получение списка номеров групп выделенных ячеек:
            selected_group_nums = []
            for selected_row_num in values["#gTable"]:
                if (selected_group_nums == [] or
                    group_nums[selected_row_num] != selected_group_nums[-1]):
                    selected_group_nums.append(group_nums[selected_row_num])
            # Получение списка номеров строк содержащих группы выделенных ячеек:
            select_rows = []
            for selected_group_num in selected_group_nums:
                for row_num in range(len(group_nums)):
                    if selected_group_num == group_nums[row_num]:
                        select_rows.append(row_num)
            window["#gTable"].Update( select_rows=select_rows )

    elif event == GT("ProcSelected"):
        if values["#gTable"] == []:
            sg.Popup(GT("SelectTableRow"), title=GT("Error"), keep_on_top=True)
        else:
            ProcData = []
            for selected_row_nums in values["#gTable"]:
                ProcData.append(TabData[selected_row_nums])
            if process(ProcData, values["#gAlgorithms"]) == 1: # функция передаёт данные конкретному плагину
                sg.Popup(GT("InternalError"), title=GT("Error"), keep_on_top=True)

    elif event == GT("ProcAll"):
        ProcData = window["#gTable"].Get()
        print(ProcData)
        if ProcData == "":
            sg.Popup(GT("EmptyTable"), title=GT("Error"), keep_on_top=True)
        else:
            if len(ProcData[0]) > 2:
                for row in ProcData:
                    del row[0]
            print(ProcData)
            # if values["#gAlgorithms"] in plugins:
            if process(ProcData, values["#gAlgorithms"]) == 1: # функция передаёт данные конкретному плагину
                sg.Popup(GT("InternalError"), title=GT("Error"), keep_on_top=True)



window.close()
