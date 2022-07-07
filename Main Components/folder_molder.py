#---------------------------------------------------------------------------------------------------------------------#
# folder_molder.py
#
# This module makes folders from Team Shop emails. It does this by referencing the TeamShopDB for accurate folder names.
#
#---------------------------------------------------------------------------------------------------------------------#

import apsw
import csv
import os
import re
import win32com.client
from configparser import ConfigParser
from pathlib import Path

# This block opens the configuration file and retrieves the database and working folder paths.

config_path = Path(__file__).parent.absolute().joinpath('config.ini')
config = ConfigParser()
config.read(config_path)

database_path = config['Folder Paths']['database_path']
working_folder_path = config['Folder Paths']['working_folder_path']

# Here, the database connection is opened, as is Outlook, and email folders are defined for later use.

connection = apsw.Connection(database_path)
cursor = connection.cursor()

outlook = win32com.client.Dispatch('Outlook.Application')
mapi = outlook.GetNamespace('MAPI')

inbox = mapi.GetDefaultFolder(6).Folders['Postorders']
winbox = mapi.GetDefaultFolder(6).Folders['Postorders, Working']
failbox = mapi.GetDefaultFolder(6).Folders['Postorders, Failed']
trashbox = mapi.Folders.Item(1).Folders['Archive']

# This gets a list of all emails in the Postorders folder and sorts them in chronological order, oldest-first.

messages = inbox.Items
messages.Sort("[ReceivedTime]", Descending=False)

# This terminates the module if there are no emails.

try:
    last_message = messages[0]
except IndexError:
    print('No emails!')
    exit()

count_file = 0
database_tuple = 0

# This creates a temporary copy of the count file so it can be examined for length.
# If it has fewer than ten lines, it's blank and can be discarded along with the email.

for attachment in last_message.Attachments:
    if '.csv' in attachment.FileName:
        count_file = attachment

if count_file:
    temp_path = os.path.join(working_folder_path, '\\', count_file.FileName)
    count_file.SaveAsFile(temp_path)

    with open(temp_path, 'r', errors="ignore") as temp_count:
        reader = csv.reader(temp_count)

        line_count = 0

        for line in reader:
            line_count += 1

    if line_count < 10:
        last_message.Move(trashbox)
        os.remove(temp_path)
        exit()
    else:
        os.remove(temp_path)

# This checks the database for a matching store id.
# If it doesn't find one, the module is terminated and the email is put aside for review.

store_id_match = re.search(r'\d\d\d\d\d\d\d\d\d\d', last_message.Subject)
store_id = store_id_match[0]

cursor.execute('SELECT * FROM team_name_index WHERE store_id=?', (store_id,))
database_tuple = cursor.fetchone()

if not database_tuple:
    last_message.UnRead = True
    last_message.Move(failbox)
    exit()

# This creates the folder.

full_path = ''.join((working_folder_path, '\\', database_tuple[1].strip(), '\\', str(database_tuple[0])))
os.makedirs(full_path, exist_ok=True)

# This puts a copy of the count sheet in the folder.

if count_file:
    count_file.SaveAsFile(os.path.join(full_path, count_file.FileName))

# This moves the email into the working box.

last_message.UnRead = True
last_message.Move(winbox)



