# Modulo comunicazione famiglie 
# Data ultima revisione: 7 agosto 2023
#
# Versione predisposta per l'utilizzo con un config.py
#
# Documents on docutils:
# https://docutils.sourceforge.io/docs/ref/rst/directives.html#class
#


import pypdf
import json, csv
import os
import sys
import subprocess
from datetime import datetime

# Add the parent directory to path
sys.path.append('..')
import config

# Import configuration file with directory structure
basedir = config.basedir
agentdir = config.agentdir
rst2odtpath = config.rst2odtpath

# app-specific parameters
appdir = 'com_famiglie'
model_file_name = 'modello_com_famiglie.txt'

# Define the paths to be used in the app scripts

parameters_path = os.path.join(agentdir, appdir)
archive_path = os.path.join(agentdir, appdir, 'archivio') 
template_path = os.path.join(parameters_path, 'modelli')
template_file_path = os.path.join(template_path, model_file_name)
stylesheet_file_path = os.path.join(template_path, 'my_style.odt')

# subprocess_cmd:   execute multiple external commands
#                   Not used in com_famiglie_b0_1_1.py

def subprocess_cmd(command):
    process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
    proc_stdout = process.communicate()[0].strip()
    print(proc_stdout)

# bracket: takes a dictionary key: value and gives back a dictionary key: { value }
def bracket(myDict):
    resDict = {}
    for key, value in myDict.items():
        resDict['{' + key + '}'] = value
    return resDict

# extract_field_values: reads the values from the first page of the .pdf file
#                       add the current date in Italian format

def extract_field_values(pdf_path):
    # Get the current date
    current_date = datetime.now().date()

    field_values = {}
    print('The pdf_path is ' + str(pdf_path))
    with open(pdf_path, 'rb') as file:
        reader = pypdf.PdfReader(file)

        # Only the first page of the form is read
        page_num = 0
        page = reader.pages[page_num]
        page_text = page.extract_text()

        for annot in page['/Annots']:
            obj = annot.get_object()
            
            if '/T' in obj:
                field_name = obj['/T']
                field_value = obj['/V']
                field_values[field_name] = field_value
                 
        # insert the current date
        italian_date = current_date.strftime("%d/%m/%Y")
        field_values['data'] = italian_date
        
    return field_values

# substitute_placeholders_in_rst: substitute placeholder in a reStructuredText file 
#                                 and return the out file

def substitute_placeholders_in_rst(file_path, out_file, placeholder_dict):
    with open(file_path, 'r', encoding="utf8") as file:
        content = file.read()

        for placeholder, value in placeholder_dict.items():
            content = content.replace(placeholder, value)

    out_file_tmp = os.path.join(parameters_path, 'temp.txt')

    with open(out_file_tmp, 'w', encoding='utf-8') as file:
        file.write(content)
    
    command = f"python {rst2odtpath} --stylesheet={stylesheet_file_path} --table-border-thickness=0 {out_file_tmp} > {out_file}"

    completed_process = subprocess.run(command, shell=True, text=True, capture_output=True)
    print(completed_process.stdout)
    print(completed_process.stderr)
    print(command)
    os.system(command)

# registra: inserts a row into file_path storing the new data on the register file

def registra(dati, file_path):

    print(file_path)
    
    # Step 1: Read the CSV file and get the last value in the first cell
    
    with open(file_path, 'r', newline='') as csv_file:
        reader = csv.reader(csv_file)
        last_row = None
        for row in reader:
            last_row = row

        if last_row:
            last_value = int(last_row[0]) + 1
        else:
            last_value = 1

        with open(file_path, 'a', newline='') as csv_file:
            writer = csv.writer(csv_file)
            out_data = dati.insert(0,last_value)
            writer.writerow(dati)  # Replace with your desired data

        return last_value

# testo_accompagnamento: generates the content of the email communication

def testo_accompagnamento(fields):
    
    rows = [f"{key}: {value}" for key, value in fields.items()]
    result_string = "\n".join(rows)
    return result_string

# crea_comunicazione: creates a new communication with ID number num and stores the communication
#                     filled with the data in ph_dicti n the out_file

def crea_comunicazione(num, ph_dict):
    
    ph_dict['{num}'] = str(num)
    out_file = os.path.join(archive_path, 'com_famiglia_' + str(num) + '.odt')

    substitute_placeholders_in_rst(template_file_path, out_file, ph_dict)
    return out_file

# processaItanza: process the request contained in pdf_file

def processaIstanza(pdf_file):

    # Import authorized users
    with open(os.path.join(parameters_path, 'auth_users.json'), 'r') as file:
        auth_users_list = json.load(file)

    # Import common parameters
    with open(os.path.join(parameters_path, 'commons_pars.json'), 'r') as file:
        common_pars_dict = json.load(file)


    # For compatibility with old versions of the script
    anno_scolastico = common_pars_dict['anno_scolastico']
    register_file_name = os.path.join(parameters_path, common_pars_dict['register_name'])
    footer = common_pars_dict['footer']
    current_date = datetime.now().date()

    res_dict = {}

    # now we process the request
    fields = extract_field_values(pdf_file)
    if fields:
        receiver_email = fields['email']

        if receiver_email not in auth_users_list:
            print('Utente ' + receiver_email + ' non autorizzato alla richiesta.\n\n')
        else:
            num_diff = registra(list(fields.values()), register_file_name)
            subject = 'Registrazione comunicazione alla famiglia n.' + str(num_diff) + ' a.s. ' + anno_scolastico
            body = '''Registrazione della comunicazione alla famiglia effettuata
                    con numero progressivo: ''' + str(num_diff) + '''\n
                   \n con i seguenti dati:\n''' + testo_accompagnamento(fields) + footer

            print('\n\nAnalisi del file:',pdf_file)
            print('\n\nNumero diffida:',num_diff)
            
            file_comunicazione = crea_comunicazione(num_diff, bracket(fields))
            attch_list = []
            attch_list.append(file_comunicazione)
            print(attch_list)

            res_dict['destinatario'] = receiver_email
            res_dict['subject'] = subject
            res_dict['body'] = body
            res_dict['attachments'] = attch_list
            print(res_dict)

    else:
        print('Non sono presenti campi nella richiesta.\n\n')
    
    return res_dict