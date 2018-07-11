# dependencies
from imapclient import IMAPClient
import email
import secrets
import os
import csv
import re

parsedData = []
tweets = []

def login_and_write():
    # connects to IMAP server 
    server = IMAPClient('imap.gmail.com', use_uid=True)
    server.login(secrets.email, secrets.pwd)
    b'[CAPABILITY IMAP4rev1 LITERAL+ SASL-IR [...] LIST-STATUS QUOTA] Logged in'

    # Selects inbox as target
    select_info = server.select_folder('INBOX')

    # Searches inbox for emails from desired address
    messages = server.search(['FROM', secrets.target_email])

    # opens csv
    info_sheet = open('fire-dept-twitter.csv', 'w')    

    # if there's no header, write a header
    if os.stat('fire-dept-twitter.csv').st_size == 0:
        info_sheet.write("Address, City, Type of Incident \n") 

    # loops through messages to fetch text content
    # reassigns the parsedData list each iteration to prevent duplicates
    for msgid, data in server.fetch(messages, ['BODY[TEXT]']).items():
        envelope = data[b'BODY[TEXT]']
        target = str(envelope.decode()[14:])
        parsedData = target.split(";")
        parsedData.pop(0)
        # The emails typically contain special characters
        # using regex to remove them
        new_string1 = re.sub('[^A-Za-z0-9 ]+', '', parsedData[0])
        new_string2 = re.sub('[^A-Za-z0-9 ]+', '', parsedData[1])
        new_string3 = re.sub('[^A-Za-z0-9 ]+', '', parsedData[2])
        info_sheet.write(new_string1 + ",")
        info_sheet.write(new_string2 + ",")
        info_sheet.write(new_string3 + "\n")

    logout(server)

# Logs out of remote session
def logout(server):
    server.logout()
    b'Logging out'

# calls login function
login_and_write()