# ---------------------------------------------------------------------------------------------------------------------#
# log_borg.py
#
# This module logs Wilson Team Shop orders in the Heat Transfer Inventory Filemaker App so that I don't have to do it.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import pyodbc
import re
from datetime import datetime
from dotenv import load_dotenv


def main(folder_path):
    load_dotenv()

    # These variables keep track of what elements are included in the outbound order, the total quantities thereof,
    # and the "in-stock" status.

    has_heat_transfers = 'No'
    has_player_numbers = 'No'
    heat_transfer_total = 0
    player_number_total = 0
    film_total = 0
    in_stock = 'Yes'

    # This block gets the Team Name.

    untrimmed_team_name = folder_path.split(os.path.sep)[-2]
    name_trimmer_regex = (re.search(r'\s\d\d', untrimmed_team_name))
    trimmer_string = name_trimmer_regex.group(0)
    team_name = untrimmed_team_name.split(trimmer_string)[0]

    # This block gets the Store Number.

    store_number = folder_path.split(os.path.sep)[-1]

    csv_path = ''

    for file in os.listdir(folder_path):
        if file.endswith('.csv'):
            csv_path = os.path.join(folder_path, file)
            break

    if csv_path == '':
        print('No count found!')
        return

    # This block sets the Order Type based on the letter code in the folder name.

    if 'WP' in folder_path:
        order_type = 'WTS'
    else:
        order_type = 'Normal'

    # This opens the count .CSV and checks it for various indicators as described below.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # This checks the count sheet for evidence of heat transfers and provides the total number of HTAs.

        for line_x in reader:
            if 'Heat Transfer' in str(line_x):
                has_heat_transfers = 'Yes'
                for line_y in reader:
                    if 'Total' in str(line_y):
                        heat_transfer_total = line_y[1]
                        break
                break

        # If there are heat transfers, this checks the count sheet for evidence of player numbers.

        csv_count.seek(0)

        if has_heat_transfers == 'Yes':
            for line in reader:
                if 'NUMBERS' in str(line):
                    has_player_numbers = 'Yes'
                    break

        # If there are player numbers, this provides the total number of Number HTAs.

        csv_count.seek(0)

        if has_player_numbers == 'Yes':
            for line in reader:
                if '# Total' in str(line):
                    player_number_total = line[1]
                    break

        # This checks the count sheet for evidence of film and provides the total number of items.

        csv_count.seek(0)

        for line_x in reader:
            if 'Film' in str(line_x):
                for line_y in reader:
                    if 'Total' in str(line_y):
                        film_total = line_y[1]
                        break
                break

    # This block checks for an order page in the folder in order to check if the order is covered by stock.

    for file in os.listdir(folder_path):
        if 'MLOrder' in str(file):
            in_stock = 'No'
            break

    # This block builds the "HTA Visual Link" string, making sure to use the correct year for the "20XX Orders" folder.

    order_year_regex = (re.search(r'\s\d\d[ABCF]', folder_path))

    if order_year_regex:
        order_year = '20' + order_year_regex.group(0)[1:3]
    else:
        order_year = datetime.now().year

    hta_visual_link = 'R:\\Transfers Heat others\\Art To Vendors\\' + order_year + ' Orders\\' + \
                      folder_path.split(os.path.sep)[-2] + '\\' + folder_path.split(os.path.sep)[-1]

    # Now that the data has been gathered, this block inserts the record into the FileMaker database.
    #
    # This has some mildly-onerous prerequisites. First, you'll need appropriate ODBC drivers for your version of
    # FileMaker. These can be downloaded from Claris' website. (Also, I have them, so if I'm still around just ask and
    # I'll hook you up.) You might need to do this on a personal computer and move the files over, but I didn't have any
    # trouble installing the drivers once I had them. If you do, bug IT about it.
    #
    # Run both "FMODBC_Installer_Win64" and "FMODBC_Installer_Win32".
    #
    # Next, you'll need to ask a FileMaker admin (at time of writing, Trey) for the "fmxdbc" permission.
    # I used to have the right to do this, but it has apparently been revoked at some point between me granting myself
    # this power and now. Luckily, the power remains, even if I am no longer its broker.
    #
    # Finally, you will need to create a DSN using the ODBC Data Source Administrator (64-bit), and potentially an
    # identically-named one using the 32-bit version. I have mixed information on the usefulness of the 32-bit copy and
    # figured "better safe than sorry", but it may not be needed.
    #
    # Use the following settings (skipped lines are times to click "Next >"):
    #
    # Name: Heat Transfer Inventory
    # Description: [Blank/Optional]
    #
    # Host: 10.7.2.137
    # Remaining Options: Optional
    #
    # Database: Heat Transfer Inventory
    # Remaining Options: Optional
    #
    # Your FileMaker User ID and Password should be saved in your environment variables as "HTI_UID" and "HTI_PWD".
    # For example:
    #
    # HTI_UID=sellickt
    # HTI_PWD=moustache

    connection_list = ['DSN=Heat Transfer Inventory;Database=Heat Transfer Inventory;UID=',
                       os.getenv('HTI_UID'),
                       ';PWD=',
                       os.getenv('HTI_PWD')]

    connection = pyodbc.connect(''.join(connection_list), autocommit=True)
    cursor = connection.cursor()

    cursor.execute('''INSERT INTO Serigraphy (TeamName, StoreNum, OrderType, StoreStatus, HasNums, TotalNUMHTA, 
                        TotalArtHTA, TotalFilmHTA, StockStatus, VisualLink) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                   team_name, store_number, order_type, 'N/A', has_player_numbers, player_number_total,
                   heat_transfer_total, film_total, in_stock, hta_visual_link)
    print('Record created!')
