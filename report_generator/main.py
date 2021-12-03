from pprint import pprint
import pdf2docx
import docx2pdf
import textract
import time
import re
import os
from os import listdir
from os.path import isfile, join
from subprocess import check_output
from pdf2docx import Converter
import argparse
import time
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
expected1 =  set(['inspectiondate', 'address', 'buildingaddress', 'owner', 'bin', 'blocklot', 'propertytype', 'contactperson', 'roof', 'walls', 'client'])
expected2 = set(['inspectiondate', 'contactperson', 'blocklot', 'propertytype', 'scopework', 'sq', 'buildingaddress', 'owneraddress', 'lin', 'bin', 'streetname', 'a', 'owner', 'd', 'tabledata', 'b', 'client', 'c'])
logger = None


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

def parse_data_1a(file_path):
    results = {}
    data = initial_data_read(file_path)
    if data is None:
        return None

    # get data
    results["inspectiondate"] = get("investigation on\n","\n",data).split(" ")[0]
    results["owneraddress"] = get("Address ","\n",data).replace("\n"," ")
    results["buildingaddress"] = get("Premise No. ","\n",data) + " " + get("Street Name ","\n",data) + " " + get("City ","\n",data) + " " + get("State ","\n",data) + " "+ get("Zip ","\n",data)
    results["owner"] = get("Building Owner ","\n",data).replace("\n"," ")
    results["bin"] = get("BIN ","\n",data)
    results["blocklot"] = get("Block ","\n",data) + "/" + get("Lot ","\n",data)
    results["propertytype"] = get("Type of Facility ","\n",data).split(" ")[0]
    results["contactperson"] = get("Contact Person ","\n",data) + " " + get("Tel. # ","\n",data)
    results["scopework"] = get("Scope of Work ","\n8.",data).replace("\n", " ") # *** (.7)
    results["streetname"] =   get("Street Name ", "\n", data).replace("\n"," ")

    # clean data
    results["owner"] = results["owner"].replace("\n", " ")
    results["contactperson"] = results["contactperson"].replace("\n", " ")
    if "\n\n" in results["propertytype"]:
        results["propertytype"] = results["propertytype"].split("\n\n")[0]

    if results["propertytype"] == "Residence":
        results["propertytype"] == "Residential"
    if " " in results["bin"]:
        results["bin"] = results["bin"].split(" ")[0]

    results["buildingaddress"] = results["buildingaddress"].replace("\n", " ")
    if "Contact Person" in results["buildingaddress"]:
        results["buildingaddress"] = results["buildingaddress"].split("Contact")[0]
    results["contactperson"] = results["contactperson"].replace("\n", " ")

    t = Converter(file_path).extract_tables()
    for ii in t:
        for i in ii:
            if len(i) > 6:
                if 'Roof' in i[2]:
                    results["roof"] = str(i[6]) + "sq ft"
                if 'Wall' in i[2] or 'Siding' in i[2]:
                    results["wall"]  = str(i[6]) + "sq ft"

    out("\t\tprocessed A!!")
    return results




def fill_templateA(results):
    load_template.update_temp("owneraddress", results["owneraddress"])
    load_template.update_temp("buildingaddress", results["buildingaddress"])
    load_template.update_temp("inspectiondate", results["inspectiondate"])

    load_template.update_temp("client", results["client"])
    load_template.update_temp("owner", results["owner"])
    load_template.update_temp("bin", results["bin"])
    load_template.update_temp("blocklot", results["blocklot"])
    load_template.update_temp("propertytype", results["propertytype"])
    load_template.update_temp("contactperson", results["contactperson"])
    load_template.update_temp("scopeofwork", results["scopework"])

    load_template.update_temp("roof", results["roof"])
    load_template.update_temp("wall", results["wall"])

"""

=====================================================================================

"""


def parse_data_2a(file_path):
    results = {}
    data = initial_data_read(file_path)

    # get data
    try:
        results["inspectiondate"] = get("investigation on\n", "\n", data,returnval=False).split(" ")[0]
    except Exception as e:
        return None
    results["owneraddress"] = get("Address ", "\n", data).replace("\n", " ")
    results["buildingaddress"] = get("Premise No. ", "\n", data) + " " + get("Street Name ", "\n", data) + " " + get(
        "City ", "\n", data) + " " + get("State ", "\n", data) + " " + get("Zip ", "\n", data)
    results["owner"] = get("Building Owner ", "\n", data).replace("\n", " ")
    results["bin"] = get("BIN ", "\n", data)
    results["blocklot"] = get("Block ", "\n", data) + "/" + get("Lot ", "\n", data)
    results["propertytype"] = get("Type of Facility ", "\n", data).split(" ")[0]
    results["contactperson"] = get("Contact Person ", "\n", data) + " " + get("Tel. # ", "\n", data)
    results["scopework"] = get("Scope of Work ", "\n8.", data).replace("\n", " ")  # *** (.7)
    results["streetname"] = get("Street Name ", "\n", data).replace("\n", " ")

    # clean data
    results["owner"] = results["owner"].replace("\n", " ")
    results["contactperson"] = results["contactperson"].replace("\n", " ")
    if "\n\n" in results["propertytype"]:
        results["propertytype"] = results["propertytype"].split("\n\n")[0]

    if results["propertytype"] == "Residence":
        results["propertytype"] == "Residential"
    if " " in results["bin"]:
        results["bin"] = results["bin"].split(" ")[0]

    results["buildingaddress"] = results["buildingaddress"].replace("\n", " ")
    if "Contact Person" in results["buildingaddress"]:
        results["buildingaddress"] = results["buildingaddress"].split("Contact")[0]
    results["contactperson"] = results["contactperson"].replace("\n", " ")

    results["a"] = ""
    results["b"] = ""
    results["c"] = ""
    results["d"] = ""
    d = data.replace("\n","").replace(" ","")
    if "Xa.the" in d:
        results["a"] = "X"
    if "Xb.premise" in d:
        results["b"] = "X"
    if "Xc.asbestos" in d:
        results["c"] = "X"
    if "Xd.entire" in d:
        results["d"] = "X"

    results['sq'] = get("Specifyamount", "linearft", d).replace(":","").split("sq.ft")[1]

    results['lin'] = get("Specifyamount", "linearft", d).replace(":","").split("sq.ft")[0]

    # get table
    t = Converter(file_path).extract_tables()
    table_data = []
    for i in t:
        if "I have advised" in i[0][0] or "The investigator" in i[0][0] or "NAME" in i[0][0] or i[0][0] == '' or i[0][0] == 'X':
            print(".")
        else:
            for rr in i:
                if len(rr) == 10:
                    table_data.append(rr)
    results["tabledata"] = table_data
    out("\t\tprocessed A!!")
    return results


def fill_templateB(results):
    load_template.update_temp("owneraddress", results["owneraddress"])
    load_template.update_temp("buildingaddress", results["buildingaddress"])
    load_template.update_temp("inspectiondate", results["inspectiondate"])

    load_template.update_temp("client", results["client"])
    load_template.update_temp("owner", results["owner"])
    load_template.update_temp("bin", results["bin"])
    load_template.update_temp("blocklot", results["blocklot"])
    load_template.update_temp("propertytype", results["propertytype"])
    load_template.update_temp("contactperson", results["contactperson"])
    load_template.update_temp("scopeofwork", results["scopework"])

    load_template.update_temp("A", results["a"])
    load_template.update_temp("B", results["b"])
    load_template.update_temp("C", results["c"])
    load_template.update_temp("D", results["d"])
    load_template.update_temp("squareft", results["sq"])
    load_template.update_temp("linearft", results["lin"])

    load_template.update_table_three_columns("table1", results["tabledata"])
    load_template.update_table("table2", results["tabledata"])
    load_template.update_table_yes("table3", results["tabledata"])
    load_template.update_table_three_columns("table4", results["tabledata"])
    load_template.update_table("table5", results["tabledata"])


def parse_data_b(file_path):
    results = {}

    data = initial_data_read(file_path)

    if data is None:
        return None

    # extract data
    try:
        t = re.search(r'Bill To\n.*?\n', data, re.DOTALL).group()
        t = t.replace("Bill To\n", "").replace('\n\nPO', "")
    except Exception as e:
        t = re.search(r'Bill To.*?\n', data, re.DOTALL).group()
        t = t.replace("Bill To", "").replace('\n', "")

    # clean data
    results["client"] = t.replace("\n", " ")
    out("\t\tprocessed B!!")
    return results


def out(t, printt=True):
    global logger
    logger.write(t + "\n")
    if printt:
        print(t)

def load_files():
    """
    loads all directories based on datadir and curr_dir
    :return:
    """
    allworkingdirectories = {}
    for dir in listdir(DATA_DIR):
        curr_dir = join(DATA_DIR, dir)
        allworkingdirectories[curr_dir] = []
        logger.write("\n" + curr_dir)
        try:
            for f in listdir(curr_dir):
                f = join(curr_dir, f)
                if isfile(f):
                    allworkingdirectories[curr_dir].append(f)
        except Exception as e:
            pass
    return allworkingdirectories



def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
def master_parser1(files, debug=True):
    file_interest = None

    # exponential big O but not expecting large # of files
    for file in files:
        file_ = file.split("\\")[-1].replace(" ", "").lower().replace("acp5", "").replace("asp5", "").replace(".", "").replace("avenue", "ave")

        if "stpdf" in file_:
            for file2 in files:
                file2_ = file2.split("\\")[-1].replace(" ", "").lower().replace("acp5", "").replace("asp5", "").replace(".", "").replace("avenue", "ave")
                if file2_ == file_:
                    continue
                score =  similar(file_,file2_)

                if score > .80:
                    file_interest = file2
                    # print("\n\t{}\n\t{}".format(file2_, str(score)))
                    break
        if file_interest is not None:
            break


    if file_interest is not None:
        print("\t\tparser1 ran on {}".format(file_interest))
        if debug:
            if MODE == 1:
                r = parse_data_1a(file_interest)
            else:
                r = parse_data_2a(file_interest)
            return r
        else:
            try:
                if MODE == 1:
                    r = parse_data_1a(file_interest)
                else:
                    r = parse_data_2a(file_interest)
                return r
            except Exception as e:
                pass


def master_parser2(files, debug=True):
    for file_path in files:
        if not debug:
            try:
                f = (file_path.split("\\")[-1].replace(".pdf", ""))

                # rule out files that arent the expected file name
                if not f[0].isupper():
                    continue
                if not f[1:].isdigit():
                    continue

                # run on file
                print("\t\tparser2 ran on {}".format(file_path))
                r = parse_data_b(file_path)
                if r is not None:
                    return r
            except Exception as e:
                pass
        else:
            f = (file_path.split("\\")[-1].replace(".pdf", ""))
            if not f[0].isupper():
                continue
            if not f[1:].isdigit():
                continue
            print("\t\tparser2 ran on {}".format(file_path))
            r = parse_data_b(file_path)
            if r is not None:
                return r


def master_writer(result, working_dir, debug=True):
    if debug:
        # print(set(result.keys()))
        # pprint(result)
        out_file = "{} {} Report.pdf".format(result['buildingaddress'], result["owner"]).replace("/", "")

        if MODE == 1:
            load_template.load_template(join(CODEBASE_DIR, "TEMPLATE1.docx"))
            fill_templateA(result)
        else:
            load_template.load_template(join(CODEBASE_DIR, "TEMPLATE2.docx"))
            fill_templateB(result)


        # write to tmp
        template_dir = join(TMP_FILES_DIR, "tmp.docx")
        load_template.write_out_filled_template(template_dir)

        # convert tmp to pdf in working dir
        docx2pdf.convert(template_dir, join(working_dir, out_file))

        # convert tmp to pdf in general working dir
        docx2pdf.convert(template_dir, join(working_dir + "\\..\\", out_file))
        return True
    else:
        try:
            out_file = "{} {} Report.pdf".format(result['buildingaddress'], result["owner"]).replace("/", "")

            if MODE == 1:
                load_template.load_template(join(CODEBASE_DIR, "TEMPLATE1.docx"))
                fill_templateA(result)
            else:
                load_template.load_template(join(CODEBASE_DIR, "TEMPLATE2.docx"))
                fill_templateB(result)

            template_dir = join(TMP_FILES_DIR, "tmp.docx")
            load_template.write_out_filled_template(template_dir)
            docx2pdf.convert(template_dir, join(working_dir, out_file))

            return True
        except Exception as e:
            # fill template
            killword()
            out("    issue with filling template and outputing report!! ({})".format(str(e)))
        return False
def main(debug=True):
    global dir
    global outdir
    global datadir
    global VERSION
    global logger
    global load_template
    global MODE

    # load template
    out("1. starting program:\n")
    tmp = join(os.path.dirname(CODEBASE_DIR), "TEMPLATE{}.docx".format(str(MODE)))

    out("2. template loading: {} \n".format(tmp))
    if MODE == 1:
        import load_template as load_template
    else:
        import load_template2 as load_template

    # load files
    logger.write("\n3. loading files stage:")
    allworkingdirectories = load_files()

    # loop each directory
    problems, success = [], []

    out("\n4. generating files stages:")
    for working_dir in allworkingdirectories.keys():

        if '\\NULL' in working_dir:
            continue
        out("\n" + working_dir)
        result = {}
        partA_done, partB_done = False, False

        r = master_parser1(allworkingdirectories[working_dir], debug=debug)
        if r is not None:
            result.update(r)
            partA_done = True

        r = master_parser2(allworkingdirectories[working_dir], debug=debug)
        if r is not None:
            result.update(r)
            partB_done = True

        # QA HERE for results
        if not partA_done or not partB_done:
            out("\t\tissue with expected files processed!!")
            problems.append(working_dir)
            continue

        # generate out_file
        report_successfully = master_writer(result,working_dir, debug=debug)
        if report_successfully:
            success.append([working_dir, DONE_DIR +"\\"+ working_dir.split("\\")[-1]])
        else:
            problems.append(working_dir)
            out("\t\treport NOT generated.")
        print('\n__________________________________________________________________________________________________________________________________________________________\n')

    out("\nfinished with: ")
    for p in success:
        try:
            os.mkdir(p[1])
            check_output("Xcopy \"{}\" \"{}\" /e /h".format(p[0], p[1]), shell=True)
            out("    " + str(p[1]))
        except Exception as e:
            out("    *FAILED* from {} to {} {}".format(p[0], p[1], e))

    out("\nissues with: ")
    for p in problems:
        out("    " + str(p))

import time
if __name__ == '__main__':
    global MODE
    parser = argparse.ArgumentParser()
    parser.add_argument('-mode', dest="mode", default=-1, type=int,help='')
    parser.add_argument('-debug', dest="debug", default=0, type=int,help='')
    results = parser.parse_args()

    MODE = results.mode
    debug = results.debug == 1


    # gather file directories
    HOME_DIR =  os.path.dirname( os.path.dirname(os.path.abspath(__file__)))
    CODEBASE_DIR =  os.path.dirname(os.path.abspath(__file__))
    TMP_FILES_DIR = HOME_DIR +"\\tmp"
    if debug:
        TMP_FILES_DIR = CODEBASE_DIR
    DATA_DIR = join(HOME_DIR, "Files_" + str(MODE))
    DONE_DIR = join(HOME_DIR, "DONE")

    killword()
    logger = open(join(TMP_FILES_DIR, "output.log"), 'w')
    main(debug=debug)
    logger.close()
