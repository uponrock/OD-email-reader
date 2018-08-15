# Dependencies
from imapclient import IMAPClient
import email, secrets, os, csv, re, datetime, re, traceback, datetime
from datetime import timedelta, date
import pandas as pd

# Function to login to IMAP and decode data
def login_and_write():
    # Connect to imap server and start tls
    imapserver = IMAPClient(host='email.townofchapelhill.org',port=143,use_uid=True,ssl=False)
    imapserver.starttls()
    print("Connecting...")
    try: 
        # Login via secrets credentials
        imapserver.login(secrets.odsusername,secrets.odspassword)
        print("Connected.") 
    except: 
        print("Not connected.")
    b'[CAPABILITY IMAP4rev1 LITERAL+ SASL-IR [...] LIST-STATUS QUOTA] Logged in'

    # Return server
    return imapserver

# Function to fetch data from chosen mailbox folder
def etl_data(server):
    # Selects inbox as target
    server.select_folder('INBOX')

    # Var to hold datetime of last hour
    last_hour_date_time = datetime.datetime.now() - timedelta(hours = 1)
    print(last_hour_date_time)

    # Open csv
    info_sheet = open('fire_dept_raw_dispatches.csv', 'a') 
    # if there's no header, write headers
    if os.stat('fire_dept_raw_dispatches.csv').st_size == 0:
        info_sheet.write("CAD,Address,City,Type of Incident,ID,ID2\n") 
        messages = server.search(['FROM', 'Orange Co EMS Dispatch'])
        print("New messages:", messages)
    else:
        messages = server.search(['SINCE', last_hour_date_time, 'FROM', 'Orange Co EMS Dispatch'])
        print("New messages:", messages)
 
    datetimes = []
    # Fetch Envelope data which contains date received
    for msgid, datedata in server.fetch(messages, ['ENVELOPE']).items():
        date_envelope = datedata[b'ENVELOPE']
        for item in date_envelope:
            # If item in list is a datetime object
            if isinstance(item, datetime.datetime):
                # Store dates in a list to hold order
                datetimes.append(item)
    
    # loops through messages to fetch text content
    for msgid, data in server.fetch(messages, ['BODY[TEXT]']).items():
        envelope = data[b'BODY[TEXT]']
        # Decode data returned
        target = envelope.decode()
        # Use regex to find the lines we want (those containing CAD: )
        dispatch_data = re.findall("((CAD:).*)\n",target)
        # Establish var to iterate through dates list
        x = 0 
        for item in dispatch_data:
            while x < len(dispatch_data):
                # Write string of items in regex cleaned data
                info_sheet.write(str(item[0]).replace(";", ", ").rstrip() + "\n")
                # Increment
                x = x + 1
    # Close CSV being used
    info_sheet.close()
    cleanup_csv(datetimes)

    # Call logout function
    logout(server)

# Logs out of remote session
def logout(server):
    server.logout()
    b'Logging out'

# Function to clean up CSV that is created 
def cleanup_csv(dateslist): 
    print("Cleaning CSV...")
    # Create pandas dataframe from original csv
    df = pd.read_csv("fire_dept_raw_dispatches.csv")
    # Delete PII rows and drop duplicate records
    del df["ID"]
    del df["ID2"]
    df.drop_duplicates(keep='first')
    # Cleaned file 
    clean_file = open("//CHFS/Shared Documents/OpenData/datasets/staging/fire_dept_dispatches_clean.csv", "a")
    # Write headers to blank clean file
    if os.stat('//CHFS/Shared Documents/OpenData/datasets/staging/fire_dept_dispatches_clean.csv').st_size == 0:
        clean_file.write(",CAD,Address,City,Type of Incident,ID,ID2,Dates\n")
    # New dates column merged into data frame
    new_column = pd.DataFrame({"Dates": dateslist})
    df = df.merge(new_column, left_index = True, right_index = True)
    # Write dataframe to new, finalized csv
    df.to_csv(clean_file, mode='a', header=False)
    print("CSV cleaned and rewritten.")

# Main function to call other functions
def main(): 
    fire_log = open("//CHFS/Shared Documents/OpenData/datasets/logfiles/fire_dispatch_log.txt", "a")
    try:     
        # Set var to hold what is returned from long_and_write()
        exchange_mail = login_and_write()
        # Call etl_data() using the exchange imap server
        etl_data(exchange_mail)
        # Log failures and successes
        fire_log.write(str(datetime.datetime.now()) + "\n" + "Logged into Exchange IMAP and saved emails." + "\n")
    except:
        fire_log.write(str(datetime.datetime.now()) + "\n" + "There was an issue when trying to log in and save emails." + "\n" + traceback.format_exc() + "\n")

# Call main
main()
