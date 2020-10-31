# ===================== IMPORTS ============================================== #
import datetime
import json
import os
import eml_parser

# ===================== CONSTANTS ============================================ #
DIR_EMAIL = "E-mails/"
DIR_CVS   = "Curriculos/"
EXT_EMAIL = ".eml"

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
    with open(eml_file, 'rb') as fhdl:
        raw_email = fhdl.read()
    ep = eml_parser.EmlParser()
    parsed_eml = ep.decode_email_bytes(raw_email)
    return parsed_eml

# ===================== MAIN SCRIPT ========================================== #
list_email_files = list_files_by_extension(DIR_EMAIL, EXT_EMAIL)
for idx in range(len(list_email_files)):
    email_dict = email_to_dictionary(list_email_files[idx])
    print(email_dict["header"]["header"]["from"][0])
    print(email_dict["attachment"][0]["filename"])
    
exit()
x = json.dumps(parsed_eml, default=json_serial)

json_file = open('teste.json','w')
json_file.write(x)
json_file.close()

print(email_dict["header"]["header"]["from"][0])
print(email_dict["attachment"][0]["filename"])

