# ===================== IMPORTS ============================================== #
import datetime
import json
import os
import eml_parser
import shutil

# ===================== CONSTANTS ============================================ #
DIR_EMAIL = "E-mails/"
DIR_CVS   = "Curriculos/"
DIR_JSON  = "Json/"
EXT_EMAIL = ".eml"
EXT_JSON  = ".json"
FILE_CANDIDATES = "lista.tsv"

# ===================== AUXILIAR FUNCTIONS =================================== #
def json_serial(obj):
    serial = ""
    if isinstance(obj, datetime.datetime):
        serial = obj.isoformat()
    return serial
    
def open_creating_dirs(path, mode):
    if not os.path.isdir(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    fp_itemx = open(path, mode)
    return fp_itemx

def list_files_by_extension(directory, extension):
    output_list = []
    for file in os.listdir(directory):
        if file.endswith(extension):
            output_list.append(os.path.join(directory, file))
    return output_list
    
def email_to_dictionary(eml_file):
    with open(eml_file, "rb") as fhdl:
        raw_email = fhdl.read()
    ep = eml_parser.EmlParser()
    parsed_eml = ep.decode_email_bytes(raw_email)
    return parsed_eml
    
def break_sender(str_name_email):
    return tuple(str_name_email[:-1].split(" <"))

# ===================== MAIN SCRIPT ========================================== #
if not os.path.isdir(os.path.dirname(DIR_CVS)):
        os.makedirs(os.path.dirname(DIR_CVS))
list_email_files = list_files_by_extension(DIR_EMAIL, EXT_EMAIL)
for idx in range(len(list_email_files)):
    # Parses to variables
    idx_str = str(idx + 1).zfill(3)
    email_file = list_email_files[idx]
    email_dict = email_to_dictionary(email_file)
    candidate = {}
    candidate["name"] = break_sender(email_dict["header"]["header"]["from"][0])[0]
    candidate["email_address"] = break_sender(email_dict["header"]["header"]["from"][0])[1]
    candidate["cv_filename"] = email_dict["attachment"][0]["filename"]
    # Json file for debug
    fp = open_creating_dirs(DIR_JSON + idx_str + EXT_JSON, "w")
    json_str = json.dumps(email_dict, default=json_serial)
    fp.write(json_str)
    fp.close()
    # Process or warn about some missing file
    cv_extension = "." + candidate["cv_filename"].split(".")[-1]
    path_to_received = DIR_EMAIL + candidate["cv_filename"]
    path_to_organized = DIR_CVS + idx_str + " - " + candidate["name"] + cv_extension
    if os.path.exists(path_to_received):
        shutil.copyfile(path_to_received, path_to_organized)
    else:
        print(idx_str)
        print(email_file)
        print(candidate)
        print()
