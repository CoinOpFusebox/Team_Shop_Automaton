# ---------------------------------------------------------------------------------------------------------------------#
# draft_counter.py
#
# This module counts the number of orders in the Drafts folder.
#
# ---------------------------------------------------------------------------------------------------------------------#
import win32com.client
from configparser import ConfigParser
from pathlib import Path
from time import sleep


def main():
    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)
    primary_mailbox = str(config['Outbot Options']['primary_mailbox'])

    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    draftbox = mapi.Folders[primary_mailbox].Folders['Drafts']

    count = 0

    messages = draftbox.Items
    messages.Sort("[ReceivedTime]", Descending=False)

    for item in messages:
        message_name = item.Subject

        if 'FILM' not in message_name and 'RDX' not in message_name:
            count += 1

    print(count, 'draft orders!')


main()
