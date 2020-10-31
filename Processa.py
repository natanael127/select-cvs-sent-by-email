# ===================== IMPORTS ============================================== #
import datetime
import json
import os
import eml_parser
import shutil
import operator
import hashlib

# ===================== CONSTANTS ============================================ #
DIR_EMAIL = "E-mails/"
DIR_CVS   = "Curriculos/"
DIR_HASH  = "Hashes/"
DIR_JSON  = "Json/"
EXT_EMAIL = ".eml"
EXT_JSON  = ".json"
FILE_CANDIDATES = "lista.tsv"
MSG_ERROR = "Erro ao obter anexo"

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
    for file_name in os.listdir(directory):
        if file_name.endswith(extension):
            output_list.append(os.path.join(directory, file_name))
    return output_list
    
def email_to_dictionary(eml_file):
    with open(eml_file, "rb") as fhdl:
        raw_email = fhdl.read()
    ep = eml_parser.EmlParser()
    parsed_eml = ep.decode_email_bytes(raw_email)
    return parsed_eml
    
def break_sender(str_name_email):
    return tuple(str_name_email[:-1].split(" <"))
    
def get_extension(file_name):
    return "." + file_name.split(".")[-1]
    
def create_dir_if_missing(dir_path):
    if not os.path.isdir(os.path.dirname(dir_path)):
        os.makedirs(os.path.dirname(dir_path))
# ===================== MAIN SCRIPT ========================================== #
# Creates missing directories if does not exist
create_dir_if_missing(DIR_CVS)
create_dir_if_missing(DIR_HASH)

# Renames files to its sha256 hash
all_input_files = os.listdir(DIR_EMAIL)
for file_name in all_input_files:
    extension = get_extension(file_name)
    input_file_path = DIR_EMAIL + file_name
    with open(input_file_path, "rb") as f:
        bytes = f.read() # read entire file as bytes
    readable_hash = hashlib.sha256(bytes).hexdigest();
    output_file_path = DIR_HASH + readable_hash + extension
    shutil.copy(input_file_path, output_file_path)

# Parses e-mail files
list_email_files = list_files_by_extension(DIR_EMAIL, EXT_EMAIL)
list_candidates = []
for idx in range(len(list_email_files)):
    # Parses to variables
    email_file = list_email_files[idx]
    email_dict = email_to_dictionary(email_file)
    candidate = {}
    candidate["name"] = break_sender(email_dict["header"]["header"]["from"][0])[0]
    candidate["email_address"] = break_sender(email_dict["header"]["header"]["from"][0])[1]
    candidate["file_extension"] = get_extension(email_dict["attachment"][0]["filename"])
    candidate["file_hash"] = email_dict["attachment"][0]["hash"]["sha256"]
    list_candidates.append(candidate)

# Parses list of candidates
list_candidates = sorted(list_candidates, key=lambda k: k["name"])
tsv_string = ""
for idx in range(len(list_candidates)):
    candidate = list_candidates[idx]
    idx_str = str(idx + 1).zfill(3)
    # Append data to TSV string
    tsv_string = tsv_string + idx_str + "\t"
    tsv_string = tsv_string + candidate["name"] + "\t"
    tsv_string = tsv_string + candidate["email_address"] + "\t"
    # Process cv or warn about some missing file
    path_to_received = DIR_HASH + candidate["file_hash"] + candidate["file_extension"]
    path_to_organized = DIR_CVS + idx_str + " - " + candidate["name"] + candidate["file_extension"]
    if os.path.exists(path_to_received):
        shutil.copyfile(path_to_received, path_to_organized)
    else:
        tsv_string = tsv_string + MSG_ERROR
    tsv_string = tsv_string + "\r\n"
    
# Dump TSV string to file
fp = open(FILE_CANDIDATES, "w")
fp.write(tsv_string)
fp.close()
