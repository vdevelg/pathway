import pyexcel

def read_lang_file(LANGUAGE):
    global dictionaries
    dictionaries = {}
    file_name = r".\lang\lang.ods"
    lang_book = pyexcel.get_book(file_name=file_name)
    for sheet in lang_book:
        dictionary = {}
        for column in sheet.columns():
            if column[0] == "KEY":
                keys = column[1:]
            if column[0].startswith(LANGUAGE):
                values = column[1:]
        for i in range(len(keys)):
            dictionary[keys[i]] = values[i]
        dictionaries[sheet.name] = dictionary

def GT(key):
    global dictionaries
    dict_name = "GT" # GUI text
    dictionary = dictionaries[dict_name]
    return dictionary[key]

def rep(key):
    global dictionaries
    dict_name = "report" # report text
    dictionary = dictionaries[dict_name]
    return dictionary[key]
