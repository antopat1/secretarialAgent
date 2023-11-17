# Auxiliary functions to read emails
# Last revised / August 3, 2023

import json
import os
from os.path import basename

import pypdf, csv
import email, imaplib, smtplib, ssl
from datetime import datetime

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Directory of json files

basedir = 'C:\SaS'

jsonPath = basedir + os.path.sep + 'agent' + os.path.sep + 'emailReader'

# Import common parameters

def extract_email_from_string(s): #OK
    start_index = s.find('<')  # Find the position of the first '<' character
    end_index = s.find('>')    # Find the position of the first '>' character

    if start_index != -1 and end_index != -1:
        # Extract the substring between the '<' and '>' characters
        email = s[start_index + 1: end_index]
        return email
    else:
        return s

def strip_last_character(input_string): #OK
    # Use slicing to remove the last character
    stripped_string = input_string[:-1]
    return stripped_string

def get_domain_from_email(email_address): #OK
    # Split the email address by '@' symbol
    parts = email_address.split('@')
    
    # Extract the domain part (second part of the split)
    domain = parts[1] if len(parts) > 1 else None
    
    return domain

# Function to fetch emails based on search criteria
def fetch_emails(search_criteria): #OK

    # Import email parameters
    with open(jsonPath + os.path.sep + 'email_pars.json', 'r') as file:
        email_pars_dict = json.load(file)

    port = email_pars_dict['port']  # For SSL
    smtp_server = email_pars_dict['smtp_server']
    sender_email = email_pars_dict['sender_email']
    password = email_pars_dict['password']

    # Connect to the IMAP server address
    mail = imaplib.IMAP4_SSL('imapmail.libero.it')

    # Login to the email account
    mail.login('sasscuolacalusco@libero.it', password)
    


    # Select the mailbox you want to access (e.g., 'INBOX')
    mail.select('INBOX')
    mail.select('Spam')

    # Search for emails based on the search_criteria
    result, data = mail.search(None, search_criteria)

    # Retrieve the list of email IDs that match the search criteria
    email_ids = data[0].split()

    # Fetch and parse the emails
    emails = []
    for email_id in email_ids:
        result, data = mail.fetch(email_id, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)
        emails.append([email_id,msg])

    # Close the connection
    mail.logout()

    return emails

def email_send(receiver_email, subject, body): #OK

    # Import email parameters
    with open(jsonPath + os.path.sep + 'email_pars.json', 'r') as file:
        email_pars_dict = json.load(file)

    port = email_pars_dict['port']  # For SSL
    smtp_server = email_pars_dict['smtp_server']
    sender_email = email_pars_dict['sender_email']
    password = email_pars_dict['password']
    

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Add attachment to message and convert message to string
    #message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(smtp_server, 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
   
    
def email_send_attch(receiver_email, subject, body, attach_list): #OK

    # Import email parameters
    with open(jsonPath + os.path.sep + 'email_pars.json', 'r') as file:
        email_pars_dict = json.load(file)

    port = email_pars_dict['port']  # For SSL
    smtp_server = email_pars_dict['smtp_server']
    sender_email = email_pars_dict['sender_email']
    password = email_pars_dict['password']  

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Add the attachments
    for file_path in attach_list:
        with open(file_path, "rb") as fil:
            part = MIMEApplication(fil.read(),Name=basename(file_path))
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_path)
        message.attach(part)
        
    # Add attachment to message and convert message to string
    #message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
    

