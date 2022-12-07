
import textract
import re
import os
from difflib import SequenceMatcher
from os import listdir
from os.path import isfile, join

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def out(t, printt=True):
    if printt:
        print(t)

def killword():
    p1 = "taskkill /f /im  WINWORD.EXE  >NUL"
    os.system(p1)
    p2 = "taskkill /f /im  WINWORD.exe  >NUL"
    os.system(p2)

def get(b1, b2, data, returnval=True):
    try:
        t = re.search(r'{}.*?{}'.format(b1, b2), data, re.DOTALL).group()
        t = t.replace(b1, "").replace(b2, "")
    except Exception as e:
        if returnval:
            return " "
        else:
            raise e
    return t

def initial_data_read(file_path):
    # extract data
    try:
        data = textract.process(file_path).decode("utf-8")
        data = data.replace("\r\n", "\n").replace("\n\n", "\n")
        if data == '':
            return None
    except Exception as e:
        return None
    return data




def load_files(DATA_DIR):
    """
    loads all directories based on datadir and curr_dir
    :return:
    """
    allworkingdirectories = {}
    for dir in listdir(DATA_DIR):
        curr_dir = join(DATA_DIR, dir)
        allworkingdirectories[curr_dir] = []
        try:
            for f in listdir(curr_dir):
                f = join(curr_dir, f)
                if isfile(f):
                    allworkingdirectories[curr_dir].append(f)
        except Exception as e:
            pass
    return allworkingdirectories
