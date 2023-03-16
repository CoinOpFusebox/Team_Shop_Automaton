# ---------------------------------------------------------------------------------------------------------------------#
# preorder_data_logger.py
#
# This module collects team name and shop number data from preorder emails in a vain attempt to create order from chaos.
#
# ---------------------------------------------------------------------------------------------------------------------#

import apsw
import re
import win32com.client
from configparser import ConfigParser
from datetime import date
from pathlib import Path


def main():
    # This block opens the configuration file and retrieves the database path.
    # Incidentally, this module and the folder_molder use a small, private database (an empty copy of which is available
    # from my GitHub; I don't have a server set up) to get team names (inconsistent) to correspond with store ID numbers
    # (consistent) to allow for responsibly-named folders. This will hopefully make things easier for everyone and allow
    # the film name and number files to be accessed with a reasonable success rate while not requiring me to place them
    # myself.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)

    database_path = config['Folder Paths']['database_path']

    # Here, the database connection is opened, as is Outlook, and email folders are defined for later use.

    connection = apsw.Connection(database_path)
    cursor = connection.cursor()

    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    inbox = mapi.GetDefaultFolder(6).Folders['Preorders']
    archive = mapi.Folders.Item(3).Folders['Archive']

    # This gets a list of all emails in the Preorders folder.

    emails = inbox.Items

    # This is a regex pattern that looks for a ten-digit number.

    store_id_pattern = re.compile(r'\d\d\d\d\d\d\d\d\d\d')

    # This is the current day, for later use.

    today = date.today()

    # This loop checks each email in the folder and (hopefully) logs its information in the database. It runs in reverse
    # to avoid confusing itself. Successfully processed emails, duplicates, and emails which are of no consequence to me
    # are archived. Unsuccessfully processed emails now also archived, because in six months not one of them has been of
    # value.
    #
    # The most consistent source of accurate, consistent versions of the desired data is in the titles of the attached
    # art pages rather than the subject; subject formatting varies among senders, and they're more typo-prone than the
    # art page filenames as well.
    #
    # Some notes on filtration:
    #
    # All preorder emails which concern me have art pages attached. Since all art pages must be sent as both .JPG and
    # .AI files, any email with one or zero attachments can be safely discarded. This catches things like pants-,
    # hard-goods-, and ProFusion-only orders.
    #
    # Emails seem to count signature images as attachments, which sometimes causes trouble when just picking the first
    # attachment. Luckily, this attachment appears to be consistently named and can be skipped over if it is found.
    #
    # Some people put hyphens between words when naming the .JPG versions. I've elected to replace them automatically
    # with spaces. I could potentially catch legitimate hyphens, but it has yet to occur.

    for item in reversed(emails):
        if item.Attachments.Count <= 1:
            item.Move(archive)
            continue

        if item.Attachments[0].FileName == 'image001.png':
            attachment_name = str(item.Attachments[1])
        else:
            attachment_name = str(item.Attachments[0])

        attachment_name = attachment_name.replace('-', ' ')

        try:
            store_id_match = re.search(store_id_pattern, attachment_name)
            store_id = store_id_match[0]

            team_name_match = re.split(store_id, attachment_name)
            team_name = team_name_match[0]
        except TypeError:
            item.Move(archive)
            continue

        # This bit removes specific typos and misappellations I'm sick of seeing.

        if 'Acadmey' in team_name:
            team_name = team_name.replace('Acadmey', 'Academy')

        if 'Crusiers' in team_name:
            team_name = team_name.replace('Crusiers', 'Cruisers')

        if 'Quckers' in team_name:
            team_name = team_name.replace('Quckers', 'Quakers')

        if 'Stars Nationals' in team_name:
            team_name = team_name.replace('Stars Nationals', 'Stars National')

        if 'Gator Baseball Academy' in team_name:
            team_name = team_name.replace('Gator Baseball Academy', 'Gators Baseball Academy')

        if 'Dirt Divas' in team_name and 'Texas' not in team_name:
            team_name = team_name.replace('Dirt Divas', 'Texas Dirt Divas')

        try:
            cursor.execute('INSERT INTO team_name_index VALUES(?,?,?)', (store_id, team_name, str(today)))
            item.Move(archive)
        except apsw.ConstraintError:
            item.Move(archive)

    # This deletes all entries older than 90 days, to keep the database svelte.

    cursor.execute('DELETE FROM team_name_index WHERE julianday(\'now\') - julianday(birthday) > 90')
