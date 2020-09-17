import shutil

def copyFiles(ProcData):
    for row in ProcData:
        cur_in_path = row[0]
        cur_out_path = row[1]
        try:
            shutil.copyfile(cur_in_path, cur_out_path, follow_symlinks=True)
        except shutil.SameFileError:
            continue

def copyFiles_print(ProcData):
    print()
    print("Привет, меня зовут copyFiles")
    print("Мне передали аргументы:")
    for i in ProcData:
        print(i)
    print()
