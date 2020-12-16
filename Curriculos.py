# ===================== IMPORTS ============================================== #
import os
import eml_parser
import shutil
import hashlib
import locale
import csv

# ===================== CONSTANTS ============================================ #
DIR_EMAIL = "E-mails/"
DIR_CVS   = "Curriculos/"
EXT_EMAIL = ".eml"
FILE_CANDIDATES = "lista.csv"
MSG_ERROR = "Erro ao obter anexo"
CVS_POSSIBLE_EXT = [".pdf", ".docx", ".doc", ".odt"]
NAME_FORBBIDEN_CHARS = ["\"", "\'", "-"]

# ===================== AUXILIAR FUNCTIONS =================================== #
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
# Locale settings
locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8") 

# Creates missing directories if does not exist
create_dir_if_missing(DIR_CVS)

# Creates a hash table for CVs
all_input_files = os.listdir(DIR_EMAIL)
files_hash_table = []
for k in range(len(all_input_files)):
    files_hash_table.append({})
    files_hash_table[k]["path"] = DIR_EMAIL + all_input_files[k]
    bytes = open(files_hash_table[k]["path"], "rb").read()
    files_hash_table[k]["hash"] = hashlib.sha256(bytes).hexdigest();

# Parses e-mail files
list_email_files = list_files_by_extension(DIR_EMAIL, EXT_EMAIL)
list_candidates = []
for idx_email in range(len(list_email_files)):
    # Parses to variables
    email_file = list_email_files[idx_email]
    email_dict = email_to_dictionary(email_file)
    list_attachments = email_dict["attachment"]
    # Looks for acceptable attachment
    for idx_attach in range(len(list_attachments)):
        if get_extension(list_attachments[idx_attach]["filename"]) in CVS_POSSIBLE_EXT:
            break
    # Composes candidate dictionary
    candidate = {}
    candidate["Nome"] = break_sender(email_dict["header"]["header"]["from"][0])[0]
    # Formats the name with no special characters
    for exclude_char in NAME_FORBBIDEN_CHARS:
        candidate["Nome"] = candidate["Nome"].replace(exclude_char,"")
    # Formats the name capitalized
    list_subnames = candidate["Nome"].split(" ")
    for idx_subname in range(len(list_subnames)):
        if list_subnames[idx_subname].lower() == "de":
            list_subnames[idx_subname] = list_subnames[idx_subname].lower()
        else:
            list_subnames[idx_subname] = list_subnames[idx_subname].capitalize()
    candidate["Nome"] = " ".join(list_subnames)
    # Direct fields
    candidate["E-mail"] = break_sender(email_dict["header"]["header"]["from"][0])[1]
    candidate["file_extension"] = get_extension(list_attachments[idx_attach]["filename"])
    candidate["file_hash"] = list_attachments[idx_attach]["hash"]["sha256"]
    list_candidates.append(candidate)

# Parses list of candidates
list_candidates = sorted(list_candidates, key=lambda k: locale.strxfrm(k["Nome"]))
for idx_cand in range(len(list_candidates)):
    candidate = list_candidates[idx_cand]
    candidate["Indice"] = str(idx_cand + 1).zfill(3)
    # Process cv or warn about some missing file
    index_hash = [d["hash"] for d in files_hash_table if "hash" in d].index(candidate["file_hash"])
    path_to_received = files_hash_table[index_hash]["path"]
    path_to_organized = DIR_CVS + candidate["Indice"] + " - " + candidate["Nome"] + candidate["file_extension"]
    if os.path.isfile(path_to_received):
        shutil.copyfile(path_to_received, path_to_organized)
        candidate["Observacao"] = ""
    else:
        candidate["Observacao"] = MSG_ERROR

# Dumps to a TSV
csv_keys = ["Indice", "Nome", "E-mail", "Observacao"]
with open(FILE_CANDIDATES,"w", newline="") as fp:
    dict_writer = csv.DictWriter(fp, csv_keys, delimiter=",", quotechar="\"", quoting=csv.QUOTE_NONNUMERIC, extrasaction="ignore")
    dict_writer.writeheader()
    dict_writer.writerows(list_candidates)

