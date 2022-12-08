import sys
sys.path.insert(0, '../shared/')
from shared import initial_data_read, killword, out, get, similar, load_files

import load_template
from pprint import pprint
import docx2pdf
import re
import os
from os import listdir
from os.path import isfile, join
from subprocess import check_output
from pdf2docx import Converter
import argparse


def parse_data_1a(file_path):
    results = {}
    data = initial_data_read(file_path)
    if data is None:
        return None

    # get data
    results["inspectiondate"] = get("investigation on\n","\n",data).split(" ")[0]
    results["inspectiondatetime"] = get("investigation on\n", "\n", data)
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
    print(file_path)
    t = Converter(file_path).extract_tables()

    for ii in t:
        for i in ii:
            if len(i) > 6:
                if 'Roof' in i[2]:
                    results["roof"] = str(i[6]) + "sq ft"
                if 'Wall' in i[2] or 'Siding' in i[2]:
                    if str(i[6]) != "":
                        results["wall"]  = str(i[6]) + "sq ft"
    if "roof" not in results:
        results["roof"] = "N/A"
    if "wall" not in results:
        results["wall"] = "N/A"
    out("\t\tprocessed A!!")
    return results


# ========================
def master_parser_part1(files, debug=True):
    file_interest = None

    # exponential big O but not expecting large # of files
    for file in files:
        file_ = file.split("\\")[-1].replace(" ", "").lower().replace(".", "").replace("avenue", "ave").replace("(1)", "").replace("(2)", "").replace("(3)", "")
        if "acp5pdf" in file_:
            file_interest = file
            break
    if file_interest is None:
        for file in files:
            file_ = file.split("\\")[-1].replace(" ", "").lower().replace("acp5", "").replace(".", "").replace("avenue", "ave").replace("(1)", "").replace("(2)", "").replace("(3)", "")
            if "stpdf" in file_:
                for file2 in files:
                    file2_ = file2.split("\\")[-1].replace(" ", "").lower().replace("acp5", "").replace("asp5", "").replace(".", "").replace("avenue", "ave").replace("(1)", "").replace("(2)", "").replace("(3)", "")
                    if file2_ == file_:
                        continue
                    score = similar(file_,file2_)

                    if score > .80:
                        file_interest = file2
                        # print("\n\t{}\n\t{}".format(file2_, str(score)))
                        break
            if file_interest is not None:
                break

    if file_interest is not None:
        print("\n\t\tparser1 ran on {}".format(file_interest))
        try:
            r = parse_data_1a(file_interest)
            return r
        except Exception as e:
            print(e)
            pass
def parse_data_1b(file_path):
    results = {}
    print(file_path)
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



def master_parser_part2(files, debug=True):
    for file_path in files:
        try:
            f = (file_path.split("\\")[-1].replace(".pdf", ""))

            # rule out files that arent the expected file name
            if not f[0].isupper():
                continue
            if not f[1:].isdigit():
                continue

            # run on file
            print("\t\tparser2 ran on {}".format(file_path))
            r = parse_data_1b(file_path)
            if r is not None:
                return r
        except Exception as e:
            if debug:
                raise e
            pass

# ========================

def master_writer(result, working_dir, debug=True):
    try:
        out_file = "{} {} Report.pdf".format(result['buildingaddress'], result["owner"]).replace("/", "")
        load_template.load_template(join(HOME_DIR, "TEMPLATE1.docx"))

        pprint(results)
        load_template.update_temp("owneraddress", result.get("owneraddress", ""))
        load_template.update_temp("buildingaddress", result.get("buildingaddress", ""))
        load_template.update_temp("inspectiondate", result.get("inspectiondate", ""))
        load_template.update_temp("inspectiondatetime", result.get("inspectiondatetime", ""))

        load_template.update_temp("client", result.get("client",""))
        load_template.update_temp("owner", result.get("owner", ""))
        load_template.update_temp("bin", result.get("bin", ""))
        load_template.update_temp("blocklot", result.get("blocklot", ""))
        load_template.update_temp("propertytype", result.get("propertytype", ""))
        load_template.update_temp("contactperson", result.get("contactperson", ""))
        load_template.update_temp("scopeofwork", result.get("scopework", ""))

        load_template.update_temp("roof", result.get("roof", ""))
        load_template.update_temp("wall", result.get("wall", ""))

        template_dir = join(TMP_FILES_DIR, "tmp.docx")
        load_template.write_out_filled_template(template_dir)
        docx2pdf.convert(template_dir, join(working_dir, out_file))

        return True
    except Exception as e:
        # fill template
        killword()
        out("    issue with filling template and outputing report!! ({})".format(str(e)))
        if debug:
            raise e
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
    tmp = join(os.path.dirname(CODEBASE_DIR), "TEMPLATE1.docx")

    out("2. template loading: {} \n".format(tmp))

    # load files
    allworkingdirectories = load_files(DATA_DIR)

    # loop each directory
    problems, success = [], []

    out("\n4. generating files stages:")
    for working_dir in allworkingdirectories.keys():

        if '\\NULL' in working_dir:
            continue
        out("\n" + working_dir)
        result = {}
        partA_done, partB_done = False, False

        r = master_parser_part1(allworkingdirectories[working_dir], debug=debug)
        if r is not None:
            result.update(r)
            partA_done = True

        r = master_parser_part2(allworkingdirectories[working_dir], debug=debug)
        if r is not None:
            result.update(r)
            partB_done = True

        # QA HERE for results
        if not partA_done :
            out("\t\tissue with expected files (part A) processed!!")
            problems.append(working_dir)
            continue
        if not partB_done:
            out("\t\tissue with expected files (part B) processed!!")
            problems.append(working_dir)
            continue

        # generate out_file
        report_successfully = master_writer(result,working_dir, debug=debug)
        if report_successfully:
            success.append([working_dir, DONE_DIR +"\\"+ working_dir.split("\\")[-1]])
        else:
            problems.append(working_dir)
            out("\t\treport NOT generated.")
        print('\n__________________________________________________________________________________________________\n')

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


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-debug', dest="debug", default=1, type=int,help='')
    results = parser.parse_args()

    debug = results.debug == 1

    # gather file directories
    HOME_DIR = os.path.dirname( os.path.dirname( os.path.dirname(os.path.abspath(__file__))))
    CODEBASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TMP_FILES_DIR = HOME_DIR +"\\tmp"
    if debug:
        TMP_FILES_DIR = CODEBASE_DIR
    DATA_DIR = join(HOME_DIR, "Files_1")
    DONE_DIR = join(HOME_DIR, "DONE")

    # run
    killword()
    main(debug=debug)
