# dependencies
from imapclient import IMAPClient
import email, secrets, os, csv, re, datetime, re
from datetime import timedelta
import pandas as pd

yesterday = datetime.datetime.today() - timedelta(days=1)
dispatch_data = []

def login_and_write():
    # Open csv
    info_sheet = open('fire-dept-twitter.csv', 'w') 
    # if there's no header, write a header
    if os.stat('fire-dept-twitter.csv').st_size == 0:
        info_sheet.write("CAD, Address, City, Type of Incident, ID, ID2, Date \n") 

    # Connect to imap server
    server = IMAPClient(host='email.townofchapelhill.org',port=143,use_uid=True,ssl=False)
    server.starttls()
    print("Connecting...")
    try: 
        server.login(secrets.odsuser,secrets.odspass)
        print("Connected.") 
    except: 
        print("Not connected.")
    b'[CAPABILITY IMAP4rev1 LITERAL+ SASL-IR [...] LIST-STATUS QUOTA] Logged in'

    # Selects inbox as target
    server.select_folder('INBOX')
    # Searches inbox for emails containing chfd
    messages = server.search(['TEXT', 'CHFD'])    

    datetimes = []
    for msgid, datedata in server.fetch(messages, ['ENVELOPE']).items():
        dateenvelope = datedata[b'ENVELOPE']
        for item in dateenvelope:
            if isinstance(item, datetime.datetime):
                datetimes.append(item)
    
    # loops through messages to fetch text content
    # reassigns the parsedData list each iteration to prevent duplicates
    for msgid, data in server.fetch(messages, ['BODY[TEXT]']).items():
        envelope = data[b'BODY[TEXT]']
        target = envelope.decode()
        dispatch_data = re.findall("((CAD:).*)\n",target)
        x = 0 
        for item in dispatch_data:
            while x < len(dispatch_data):
                info_sheet.write(str(item[0]).replace(";", ", ").rstrip() + ", " + str(datetimes[x]) + "\n")
                x = x + 1
    info_sheet.close()
    logout(server)

# Logs out of remote session
def logout(server):
    server.logout()
    b'Logging out'

# calls login function
login_and_write()
