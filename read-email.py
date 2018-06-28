# dependencies
from imapclient import IMAPClient
import email
import secrets

# connects to IMAP server 
server = IMAPClient('imap.gmail.com', use_uid=True)
server.login(secrets.email, secrets.pwd)
b'[CAPABILITY IMAP4rev1 LITERAL+ SASL-IR [...] LIST-STATUS QUOTA] Logged in'

# Selects inbox as target
select_info = server.select_folder('INBOX')
print('%d messages in INBOX' % select_info[b'EXISTS'])

# Searches inbox for emails from desired address
messages = server.search(['FROM', secrets.target_email])
print("%d messages from our best friend" % len(messages))

# loops through messages to fetch text content
for msgid, data in server.fetch(messages, ['BODY[TEXT]']).items():
    envelope = data[b'BODY[TEXT]']
    # this method is too specific and could cause problems
    # researching better solutions
    target = str(envelope.decode()[9697:9745])
    parsedData = target.split(";")
    parsedData.pop(-1)
    print(parsedData)
    # CSV writing functionality to go here

server.logout()
b'Logging out'