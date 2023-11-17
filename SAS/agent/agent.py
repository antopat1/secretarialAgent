# Multitask scheduler - last revised ... August 18, 2023

from croniter import croniter
from datetime import datetime
import time

from emailReader import emailReader


import os, shutil
import glob
import pypdf

# Callable modules
# Put here the list of modules processing the different forms in the office

from com_famiglie import com_famiglie

# Folder structure
basedir = 'C:\SaS' #da muovere in config.py

istanze_path = basedir + os.path.sep + os.path.sep + 'agent' + os.path.sep + 'istanze'
archivio_path = basedir + os.path.sep + os.path.sep + 'agent' + os.path.sep + 'istanze' + os.path.sep + 'archivio' # correggere automation24 per compatibilità windows-mac

# Dictionary of correspondence module_type - wrap functions (i.e. the functions inside the scheduler
# that call the functions inside the modules)

function_dict = {
    'mod_comunicazioni' : 'wrap_comunicazioni_famiglia'
}

# Wrap functions

def wrap_comunicazioni_famiglia(file_path):
    return com_famiglie.processaIstanza(file_path)


def rename_file_strip_prefix(directory, file_name):

    base_file_name = os.path.basename(file_name)
    old_file_path = os.path.join(directory, base_file_name )

    if base_file_name.startswith('inevaso_'):
        new_file_name = base_file_name[len('inevaso_'):]
        new_file_path = os.path.join(directory, new_file_name)

        os.rename(old_file_path, new_file_path)
        print(f"File '{base_file_name}' renamed to '{new_file_name}'.")
    else:
        print(f"File '{base_file_name}' does not start with 'inevaso_', no rename needed.")

# Notificatore: the function sends an email with specified files, message and subjet

def notificatore(receiver_email, subject, body, list_of_attachments):
    #
    if not list_of_attachments:
        emailReader.email_send(receiver_email, subject, body)
    else:
        emailReader.email_send_attch(receiver_email, subject, body, list_of_attachments)

# Analyzer: the function takes the form name and calls the appropriate wrap function

def analyzer(form_name, pdf_path):
        # Check if the variable form_name is a valid key in the dictionary
            # Use globals() to get the dictionary of the global symbol table
            global_symbols = globals()

            if form_name in function_dict:
                function_name = function_dict[form_name]
                print(function_name)

                if function_name in global_symbols and callable(global_symbols[function_name]):
                # Call the function
                    res_dict = global_symbols[function_name](pdf_path)
                    print('Analyzer:')
                    print(res_dict)
            else:
                print("Invalid key provided. The task cannot be processed.")
            return res_dict

def process_pdf_files(directory):
    # Get a list of all .pdf files in the given directory
    pdf_files = glob.glob(os.path.join(directory, '*.pdf'))

    for pdf_file in pdf_files:
        
        print(f"Processing: {pdf_file}")
        if 'inevaso_' in pdf_file:
        	with open(pdf_file, 'rb') as file:
        		reader = pypdf.PdfReader(file)
        		# The form has only one page
        		page_num = 0
        		page = reader.pages[page_num]
        		page_text = page.extract_text()

        		id_flag = 0

        		for model_key in function_dict.keys():
        			if model_key in page_text:
        				id_flag = 1
        				print('Identificato modello:',model_key)
        				modello_nome = model_key

        		if id_flag == 0:
        			print('Il file non è un formulario di IC Calusco.')
        			os.remove(pdf_file)

        		if id_flag == 1:
        			res_dict = analyzer(modello_nome, pdf_file)
        			if res_dict:
        				email_destinatario = res_dict['destinatario']
        				subject = res_dict['subject']
        				body = res_dict['body']
        				list_of_attachments = res_dict['attachments']
        				print('Invio a ', email_destinatario)

        				notificatore(email_destinatario, subject, body, list_of_attachments)

        	shutil.move(pdf_file, archivio_path)
        	rename_file_strip_prefix(archivio_path, pdf_file)
                
        # Rename the file form 'inevaso' to 'completato'

def main():
   
    def test_task():

        process_pdf_files(istanze_path)        

    # the script starts here
    # Fetch the unseen emails in sas@iccarvico.edu.it
    def my_task():

        # Step 1 - read unseen emails

        search_criteria = 'UNSEEN'
        emails = emailReader.fetch_emails(search_criteria)
        
        if len(emails) == 0:
            print('SaS: No email to be processed. Waiting for the next run ...')
        else:
            print('SaS: There are ' + str(len(emails)) + ' unseen emails to be processed. Let\'s have a look ... \n\n')

            if emails:

                for email_obj in emails:

                    hasAttachFlag = 0
                    
                    # get the email ID
                    email_id = email_obj[0]
                    msg = email_obj[1]
                    print('Processing email with ID:' + str(email_id) + '\n\n')
        
                    # check if the message comes from iccarvico.edu.it

                    applicant_email = emailReader.extract_email_from_string(msg['From'])
                    domain_email = str(emailReader.get_domain_from_email(applicant_email))

                    print('The email comes from ' + applicant_email + '\n\n')

                    if emailReader.get_domain_from_email(applicant_email) != 'scuolacalusco.edu.it':
                        print('Email received from outside iccalusco.edu.it / Sender: ' + msg['From'] + '\n')
                        print('The message will not be processed. \n\n')
                        continue

                    if msg.get_content_maintype() == 'multipart':


                     for part in msg.walk():

                        content_type = part.get_content_type()
                        timestamp = str(datetime.now()).replace('.','-').replace(':','-')
                        fileRef = 'inevaso_file_' + timestamp 
                        
                        # If it's a PDF attachment, save it to disk
                        
                        if content_type == 'application/pdf':
                        
                            filename = istanze_path + os.path.sep + fileRef + '.pdf'
                        
                            with open(filename, 'wb') as f:
                        
                                f.write(part.get_payload(decode=True))
                                hasAttachFlag = 1
                                                            

                    if hasAttachFlag == 0:
                        print('Email without .pdf attachment. It will be skipped.\n\n')
        
        # Step 2 - process all submitted PDF forms in istanze_path
        #          Look only at those 'inevase', after processing rename them

        process_pdf_files(istanze_path)

        pass


    # Welcome message
    print("SaS up and running!\n")

    # Define the cron-like schedule to run every 30 minutes
    cron_schedule = "*/1 * * * *"

    # Create a croniter object with the schedule
    base = datetime.now()
    cron = croniter(cron_schedule, base)

    while True:
      # Calculate the next run time
        next_run_time = cron.get_next(datetime)
        time_to_wait = (next_run_time - datetime.now()).total_seconds()
        print('SaS: next execution foreseen at ',next_run_time)
        # Wait until the next run time
        time.sleep(max(0, time_to_wait))

        # Execute the task
        my_task()
        

if __name__ == "__main__":
    main()
