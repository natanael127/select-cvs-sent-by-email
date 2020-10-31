# ===================== IMPORTS ============================================== #
import datetime
import json
import eml_parser

# ===================== AUXILIAR FUNCTIONS =================================== #
def json_serial(obj):
    if isinstance(obj, datetime.datetime):
        serial = obj.isoformat()
        return serial

# ===================== MAIN SCRIPT ========================================== #
with open('E-mails/[vagas.fw] CV para vaga estágio - Desenvolvedor em linguagem C.eml', 'rb') as fhdl:
    raw_email = fhdl.read()

ep = eml_parser.EmlParser()
parsed_eml = ep.decode_email_bytes(raw_email)
x = json.dumps(parsed_eml, default=json_serial)

json_file = open('teste.json','w')
json_file.write(x)
json_file.close()

print(parsed_eml["header"]["header"]["from"][0])
print(parsed_eml["attachment"][0]["filename"])
