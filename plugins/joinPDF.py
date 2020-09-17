import os
import subprocess

def joinPDF(argument):
    print()
    print("Привет, меня зовут joinPDF и я не реализована")
    print()

def joinPDF_print(ProcData):
    print()
    print("Привет, меня зовут joinPDF")
    print("Мне передали аргументы:")
    for i in ProcData:
        print(i)
    print()




def handlePDF(path, method, files):
    tool = "pdf24-DocTool"
    for file in files:
        command = tool + path
    os

def procFile(cmd, inFiles):

    if cmd == "split":
        flag = "-splitByPage"
    elif cmd == "join":
        flag = "-join"
    
    for i in range(len(inFiles)):
        if not inFiles[i] == "":
            inFiles[i] = "\"" + inFiles[i] + "\""
    inFiles = " ".join(inFiles).rstrip()
            
    paths = {"prog" : r"C:\Program Files (x86)\PDF24\pdf24-DocTool.exe",
             "outD" : r"C:\exper\PDF",
             "outF" : r"-",
             "inFs" : inFiles}

    command = ["{0}".format(paths['prog']),
               flag,
               #"-outputDir \"{0}\"".format(paths['outD']),
               "-outputFile \"{0}\"".format(paths['outF']),
               paths['inFs']]
    return command



 # print(procFile("split", inSplit))
 # print()
 # print(procFile("join", inJoin))


 # subprocess.call(procFile("split", inSplit))

##wd = os.getcwd()
##print(wd)
##print()
##
##for dirpath, dirnames, filenames in os.walk(wd):
##    print(dirpath, dirnames, filenames)
##    print()
##
####prog
####os.system(command)
