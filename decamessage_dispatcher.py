# ---------------------------------------------------------------------------------------------------------------------#
# decamessage_dispatcher.py
#
# This module sends ten orders' worth of emails from the Drafts folder.
#
# ---------------------------------------------------------------------------------------------------------------------#
import calendar
import win32com.client
from configparser import ConfigParser
from pathlib import Path
from time import sleep


def main():
    # This block opens Outlook and defines the relevant folders.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)
    primary_mailbox = str(config['Outbot Options']['primary_mailbox'])

    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    draftbox = mapi.Folders[primary_mailbox].Folders['Drafts']
    outbox = mapi.Folders[primary_mailbox].Folders['Outbox']

    count = 0

    # This bit closes the preview pane because attempting to send a draft that is being previewed causes an error.
    # If it was open, it will reopen at the end for the user's convenience.

    preview_pane_open = outlook.ActiveExplorer().IsPaneVisible(3)
    outlook.ActiveExplorer().ShowPane(3, False)

    while count < 10:
        messages = draftbox.Items
        messages.Sort("[ReceivedTime]", Descending=False)

        if messages:
            try:
                message_name = messages[0].Subject
            except IndexError:
                print('\nAll drafts sent!')
                return

            if 'FILM' not in message_name and 'RDX' not in message_name:
                count += 1

            messages[0].Send()

            print(message_name, 'sending...')

            while len(outbox.Items):
                sleep(1)

            print(message_name, 'sent!')
        else:
            break

    messages = draftbox.Items
    messages.Sort("[ReceivedTime]", Descending=False)

    if messages:
        try:
            message_name = messages[0].Subject
        except IndexError:
            print('\nAll drafts sent!')
            return

        if 'RDX' in message_name:
            messages[0].Send()

            print(message_name, 'sending...')

            while len(outbox.Items):
                sleep(1)

            print(message_name, 'sent!')

    if count == 10:
        print('\nBatch sent!')
    else:
        print('\nAll drafts sent!')

    if preview_pane_open:
        outlook.ActiveExplorer().ShowPane(3, True)


main()
