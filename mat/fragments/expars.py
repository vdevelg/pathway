import os
listOfStrings = ["dir", "dir", "fname", "fname", ".ftype", "fname", ".ftype", 
                 "dir", "dir", "fname", "fname", ".ftype", "fname", ".ftype",
                 "dir", "dir", "fname", "fname", ".ftype", "fname", ".ftype", 
                 "dir", "dir", "fname", "fname", ".ftype", "fname", "fname"]

def make_file_paths(dir_path, file_names, file_types, file_paths):
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


DirMarker = "dir"
FNameMarker = "fname"
FTypeMarker = ".ftype"

prev_part = ""
dir_path = ""
file_paths = []
file_names = []
file_types = []
for cur_part in listOfStrings:
    if   (prev_part, cur_part) == ("", DirMarker):
        dir_path = os.path.join( dir_path, cur_part )
    elif (prev_part, cur_part) == (DirMarker, DirMarker):
        dir_path = os.path.join( dir_path, cur_part )
    elif (prev_part, cur_part) == (DirMarker, FNameMarker):
        file_names = [cur_part]
    elif (prev_part, cur_part) == (DirMarker, FTypeMarker):
        print("!Ошибка 1 формата входного файла!")
    elif (prev_part, cur_part) == (FNameMarker, DirMarker):
        print("!Ошибка 2 формата входного файла!")
    elif (prev_part, cur_part) == (FNameMarker, FNameMarker):
        file_names.append(cur_part)
    elif (prev_part, cur_part) == (FNameMarker, FTypeMarker):
        file_types = [cur_part]
    elif (prev_part, cur_part) == (FTypeMarker, DirMarker):
        make_file_paths(dir_path, file_names, file_types, file_paths)
        file_names = []
        file_types = []
        dir_path = cur_part
    elif (prev_part, cur_part) == (FTypeMarker, FNameMarker):
        make_file_paths(dir_path, file_names, file_types, file_paths)
        file_names = [cur_part]
        file_types = []
    elif (prev_part, cur_part) == (FTypeMarker, FTypeMarker):
        file_types.append(cur_part)
    else:
        print("!Неизвестный формат входного файла!", prev_part, cur_part)
    prev_part = cur_part
if cur_part != FTypeMarker:
    print("!Ошибка 3 формата входного файла!")
else:
    make_file_paths(dir_path, file_names, file_types, file_paths)


print("Исходный список:")
for string in listOfStrings:
    print(string)

print()

print("Список путей:")
for file_path in file_paths:
    print(file_path)
