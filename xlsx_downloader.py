import imaplib
import email
import os
import json

# Load configuration from JSON file
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

username = config['username']
password = config['password']
folder = config['folder']
server = config['server']
download_path = config['download_path']

# Get the absolute path to the download folder relative to the script location
base_dir = os.path.dirname(os.path.abspath(__file__))
download_path = os.path.join(base_dir, download_path)

# Create the download folder if it doesn't exist
if not os.path.exists(download_path):
    os.makedirs(download_path)

# Making connection to IMAP server
imap = imaplib.IMAP4_SSL(server)
imap.login(username, password)
imap.select(folder)

result, data = imap.search(None, '(SUBJECT "Daily Excel File")')  # Agreed format can be placed here (could be current date)
print('The result is:', result)
print('The data is:', data[0])

for num in data[0].split():
    result, data = imap.fetch(num, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)
    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            if filename.lower().endswith('.xlsx'):  # Other checks according to the agreed name can be applied
                file_path = os.path.join(download_path, filename)
                with open(file_path, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    print(filename, 'has been downloaded to', file_path)

imap.close()
imap.logout()
