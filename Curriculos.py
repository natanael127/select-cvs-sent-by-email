# ===================== IMPORTS ============================================== #
import os
import eml_parser
import shutil
import hashlib
import locale

# ===================== CONSTANTS ============================================ #
DIR_EMAIL = "E-mails/"
DIR_CVS   = "Curriculos/"
DIR_HASH  = "Hashes/"
EXT_EMAIL = ".eml"
FILE_CANDIDATES = "lista.tsv"
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
create_dir_if_missing(DIR_HASH)

# Renames files to theirs sha256 hash
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
    candidate["name"] = break_sender(email_dict["header"]["header"]["from"][0])[0]
    # Formats the name with no special characters
    for exclude_char in NAME_FORBBIDEN_CHARS:
        candidate["name"] = candidate["name"].replace(exclude_char,"")
    # Formats the name capitalized
    list_subnames = candidate["name"].split(" ")
    for idx_subname in range(len(list_subnames)):
        if list_subnames[idx_subname].lower() == "de":
            list_subnames[idx_subname] = list_subnames[idx_subname].lower()
        else:
            list_subnames[idx_subname] = list_subnames[idx_subname].capitalize()
    candidate["name"] = " ".join(list_subnames)
    # Direct fields
    candidate["email_address"] = break_sender(email_dict["header"]["header"]["from"][0])[1]
    candidate["file_extension"] = get_extension(list_attachments[idx_attach]["filename"])
    candidate["file_hash"] = list_attachments[idx_attach]["hash"]["sha256"]
    list_candidates.append(candidate)

# Parses list of candidates
list_candidates = sorted(list_candidates, key=lambda k: locale.strxfrm(k["name"]))
tsv_string = ""
for idx_cand in range(len(list_candidates)):
    candidate = list_candidates[idx_cand]
    idx_str = str(idx_cand + 1).zfill(3)
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
    tsv_string = tsv_string + "\n"
    
# Dump TSV string to file
fp = open(FILE_CANDIDATES, "w")
fp.write(tsv_string)
fp.close()
