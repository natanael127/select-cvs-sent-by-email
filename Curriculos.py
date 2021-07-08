# ===================== IMPORTS ============================================== #
import os
import eml_parser
import locale
import csv
import base64

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
    ep = eml_parser.EmlParser(include_attachment_data=True)
    return ep.decode_email(eml_file)
    
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
    candidate["file_content"] = base64.b64decode(list_attachments[idx_attach]["raw"])
    list_candidates.append(candidate)

# Parses list of candidates and organizes documents
list_candidates = sorted(list_candidates, key=lambda k: locale.strxfrm(k["Nome"]))
for idx_cand in range(len(list_candidates)):
    candidate = list_candidates[idx_cand]
    candidate["Indice"] = str(idx_cand + 1).zfill(3)
    path_cv_doc = DIR_CVS + candidate["Indice"] + " - " + candidate["Nome"] + candidate["file_extension"]
    with open(path_cv_doc, "wb") as fp:
        fp.write(candidate["file_content"])

# Dumps to a CSV
csv_keys = ["Indice", "Nome", "E-mail"]
with open(FILE_CANDIDATES,"w", newline="") as fp:
    dict_writer = csv.DictWriter(fp, csv_keys, delimiter=",", quotechar="\"", quoting=csv.QUOTE_NONNUMERIC, extrasaction="ignore")
    dict_writer.writeheader()
    dict_writer.writerows(list_candidates)
