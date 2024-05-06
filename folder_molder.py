# ---------------------------------------------------------------------------------------------------------------------#
# folder_molder.py
#
# This module makes folders from Team Shop emails.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import pyodbc
import re
import win32com.client

from configparser import ConfigParser
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path


def main():
    # This block opens the configuration file and retrieves the working folder path.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)

    working_folder_path = config['Folder Paths']['working_folder_path']

    # This opens Outlook and defines email folders for later use.
    # TODO: Stop using the wrong archive!! WHY IS THIS SO HARD!?

    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    inbox = mapi.GetDefaultFolder(6).Folders['Postorders']
    archive = mapi.Folders.Item(3).Folders['Archive']

    # This gets a list of all emails in the Postorders folder and sorts them in chronological order, oldest-first.

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", Descending=False)

    # This terminates the module if there are no emails.

    try:
        last_message = messages[0]
    except IndexError:
        return 'No emails!'

    count_file = None

    # This creates a temporary copy of the count file so it can be examined for length.
    # If it has fewer than ten lines, it's blank and can be archived along with the email.

    for attachment in last_message.Attachments:
        if '.csv' in attachment.FileName and 'Count' in attachment.FileName:
            count_file = attachment

    if count_file:
        temp_path = os.path.join(working_folder_path, '\\', count_file.FileName)
        count_file.SaveAsFile(temp_path)

        with open(temp_path, 'r', errors="ignore") as temp_count:
            reader = csv.reader(temp_count)

            line_count = 0

            for __ in reader:
                line_count += 1

        if line_count < 10:
            last_message.Move(archive)
            os.remove(temp_path)
            print('Blank count!\n')
            return 'Blank count!'
        else:
            os.remove(temp_path)

    # This gets the store id from the email subject.

    store_id_match = re.search(r'\d\d\d\d\d\d\d\d\d\d', last_message.Subject)
    store_id = store_id_match[0]

    # This gets the team name from the WTS Status database using the store id.
    # If it doesn't find one, it uses the title of the email instead. Might as well try!

    load_dotenv()

    connection_list = ['DSN=WTS Status;Database=WTS Status;UID=',
                       os.getenv('WTSS_UID'),
                       ';PWD=',
                       os.getenv('WTSS_PWD')]

    connection = pyodbc.connect(''.join(connection_list))
    cursor = connection.cursor()

    try:
        team_name = cursor.execute('SELECT "Team Name" FROM "WTS Status" WHERE "Store ID 2"=?', store_id).fetchval()
    except TypeError:
        team_name = str(last_message.Subject).split('-')[0]

    if '  ' in team_name:
        team_name = team_name.replace('  ', ' ')

    # This creates the folders.

    year_awp_cwp = ''.join((' ', str(datetime.now().year)[-2:], 'AWP CWP'))

    full_path = ''.join((working_folder_path, '\\', team_name, year_awp_cwp, '\\', store_id))
    try:
        os.makedirs(full_path, exist_ok=True)
    except ValueError:
        full_path = full_path.replace('\x00', '')
        os.makedirs(full_path, exist_ok=True)

    # This puts a copy of the count sheet in the folder.

    if count_file:
        count_file.SaveAsFile(os.path.join(full_path, count_file.FileName))

    # This archives the email.

    last_message.Move(archive)

    return full_path
