# dependencies
from imapclient import IMAPClient
import email
import secrets
import os
import csv

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

    info_sheet = open('fire-dept-twitter.csv', 'w')    

    if os.stat('fire-dept-twitter.csv').st_size == 0:
        info_sheet.write("Address, City, Type of Incident \n")

    # loops through messages to fetch text content
    # reassigns the parsedData list each iteration to prevent duplicates
    for msgid, data in server.fetch(messages, ['BODY[TEXT]']).items():
        envelope = data[b'BODY[TEXT]']
        target = str(envelope.decode()[14:67])
        parsedData = target.split(";")
        parsedData.pop(0)
        info_sheet.write(parsedData[0] + ",")
        info_sheet.write(parsedData[1] + ",")
        info_sheet.write(parsedData[2] + "\n")

    logout(server)

# Logs out of remote session
def logout(server):
    server.logout()
    b'Logging out'

# calls login function
login_and_write()