# ---------------------------------------------------------------------------------------------------------------------#
# count_comparitron.py
#
# This module compares a list of outbound items' quantities to the corresponding inventory quantities.
# If the inventory is not sufficient to fill the order, it returns an inbound ordering list.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os

import pyodbc
import re
from dotenv import load_dotenv


def main(outbound_list):
    # This module starts by opening an ODBC connection to the FileMaker database.
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

    load_dotenv()

    connection_list = ['DSN=Heat Transfer Inventory;Database=Heat Transfer Inventory;UID=',
                       os.getenv('HTI_UID'),
                       ';PWD=',
                       os.getenv('HTI_PWD')]

    connection = pyodbc.connect(''.join(connection_list))
    cursor = connection.cursor()

    # outbound_list is a list of HTAs that have been ordered by the customer.
    # inbound_list is a list of HTAs that we don't already have in the warehouse, which must be made or ordered.

    inbound_list = []

    for i in range(len(outbound_list)):
        # This block takes the whole name of the HTA and splits it into chunks with which the database is comfortable.
        # It also gets the quantity needed from the tuple in question.

        hta_tuple = outbound_list[i]

        whole_hta_string = str(hta_tuple[0])
        whole_hta_list = re.split("HTA", whole_hta_string)

        hta_name = whole_hta_list[0]
        hta_number = 'HTA' + whole_hta_list[1]

        units_needed = int(hta_tuple[1])

        # Provided the quantity is greater than zero, this block checks the quantity in stock and compares it to the
        # quantity required. If it exceeds the requirement by fifteen or more, the item is considered covered.
        # Otherwise, it is added to the inbound_list.
        #
        # The quantity to be ordered is equal to half-again the quantity required, with a minimum order of twenty.
        #
        # Due to human error, HTA names occasionally contain unwelcome spaces. In this case, the database (or the
        # database guy, not sure) either replaces the spaces with hyphens or just omits them. Both possibilities are
        # checked for here.
        #
        # When in doubt, this module errs on the side of ordering what it doesn't understand.

        if units_needed > 0:
            try:
                units_on_hand = int(cursor.execute('SELECT "Units on Hand" FROM Inventory WHERE Name=? AND "HTA '
                                                   'Number"=?',
                                                   hta_name, hta_number).fetchval())

                if units_on_hand - units_needed < 15:
                    if units_needed <= 13:
                        units_to_order = 20
                    else:
                        units_to_order = (5 * round((units_needed * 1.5) / 5))
                    inbound_tuple = (whole_hta_string, units_to_order)
                    inbound_list.append(inbound_tuple)
            except TypeError:
                if ' ' in hta_name:
                    hta_name = hta_name.replace(' ', '-')
                    try:
                        units_on_hand = int(
                            cursor.execute('SELECT "Units on Hand" FROM Inventory WHERE Name=? AND "HTA Number"=?',
                                           hta_name, hta_number).fetchval())

                        if units_on_hand - units_needed < 15:
                            if units_needed <= 13:
                                units_to_order = 20
                            else:
                                units_to_order = (5 * round((units_needed * 1.5) / 5))
                            inbound_tuple = (whole_hta_string, units_to_order)
                            inbound_list.append(inbound_tuple)
                    except TypeError:
                        hta_name = hta_name.replace('-', '')
                        try:
                            units_on_hand = int(
                                cursor.execute('SELECT "Units on Hand" FROM Inventory WHERE Name=? AND "HTA Number"=?',
                                               hta_name, hta_number).fetchval())

                            if units_on_hand - units_needed < 15:
                                if units_needed <= 13:
                                    units_to_order = 20
                                else:
                                    units_to_order = (5 * round((units_needed * 1.5) / 5))
                                inbound_tuple = (whole_hta_string, units_to_order)
                                inbound_list.append(inbound_tuple)
                        except TypeError:
                            if units_needed <= 13:
                                units_to_order = 20
                            else:
                                units_to_order = (5 * round((units_needed * 1.5) / 5))
                            inbound_tuple = (whole_hta_string, units_to_order)
                            inbound_list.append(inbound_tuple)
                else:
                    if units_needed <= 13:
                        units_to_order = 20
                    else:
                        units_to_order = (5 * round((units_needed * 1.5) / 5))
                    inbound_tuple = (whole_hta_string, units_to_order)
                    inbound_list.append(inbound_tuple)

    if inbound_list:
        return inbound_list
    else:
        return None
