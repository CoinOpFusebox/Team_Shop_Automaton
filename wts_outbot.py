# ---------------------------------------------------------------------------------------------------------------------#
# wts_outbot.py
#
# This module prepares Wilson Team Shop order emails.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import win32com.client as win32
from configparser import ConfigParser
from pathlib import Path


def main(folder_path):
    # These variables keep track of what elements are included in the outbound order and which, if any, need to be
    # included in the inbound order.
    #
    # Some rules/assumptions:
    #
    # Heat transfers and player numbers (which are also heat transfers) are kept on hand and may already be in stock.
    # Otherwise, they need to be either made in-house or ordered, a decision beyond the purview of the art department.
    #
    # Film and embroidery are produced on demand, so if the order has them you can assume it needs them as well.
    #
    # Player numbers never appear on a garment without other heat transfers. If there are no heat transfers then there
    # are no numbers.

    has_heat_transfers = False
    needs_heat_transfers = False
    has_player_numbers = False
    needs_player_numbers = False
    has_film = False
    has_embroidery = False

    # This block finds the path for the count .CSV in the folder. If it fails to do so, the module is terminated.

    csv_path = ''

    for file in os.listdir(folder_path):
        if file.endswith('.csv'):
            csv_path = os.path.join(folder_path, file)
            break

    if csv_path == '':
        print('No count found!')
        return

    # This opens the count .CSV and checks it for various indicators as described below.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # This checks the count sheet for evidence of heat transfers.

        for line in reader:
            if 'Heat Transfer' in str(line):
                has_heat_transfers = True
                break

        # If there are heat transfers, this checks the count sheet for evidence of player numbers.

        csv_count.seek(0)

        if has_heat_transfers:
            for line in reader:
                if 'NUMBERS' in str(line):
                    has_player_numbers = True
                    break

        # If there are player numbers, this checks if they need to be ordered.

        csv_count.seek(0)

        if has_player_numbers:
            for line in reader:
                if 'ORDER' in str(line):
                    needs_player_numbers = True
                    break

        # This checks the count sheet for evidence of embroidered items.

        csv_count.seek(0)

        for line in reader:
            if 'Embroidery' in str(line):
                has_embroidery = True
                break

    # This checks the folder for evidence of film items.

    for file in os.listdir(folder_path):
        if 'Film Art Actual' in str(file):
            has_film = True
            break

    # This checks the folder for evidence of heat transfers order pages, provided the order features heat transfers.

    if has_heat_transfers:
        for file in os.listdir(folder_path):
            if 'ORDER Art Page' in str(file):
                needs_heat_transfers = True
                break

    # This block selects the type of email(s) that will be sent based on the information gathered above.

    if has_film and (needs_heat_transfers or needs_player_numbers):
        compose_combo(needs_heat_transfers, has_player_numbers, needs_player_numbers, has_embroidery, folder_path,
                      csv_path)
    elif has_film:
        compose_film(has_heat_transfers, has_player_numbers, has_embroidery, folder_path, csv_path)
    elif needs_heat_transfers or has_player_numbers:
        compose_transfer(needs_heat_transfers, has_player_numbers, needs_player_numbers, has_embroidery, folder_path,
                         csv_path)
    elif has_embroidery and not (has_heat_transfers or has_player_numbers):
        compose_embroidery(folder_path, csv_path)
    else:
        compose_stock(has_embroidery, folder_path, csv_path)


def compose_combo(needs_heat_transfers, has_player_numbers, needs_player_numbers, has_embroidery, folder_path,
                  csv_path):
    # This function is used for orders that need heat transfers or numbers AND film elements created.
    # For some reason, this situation requires the creation of two separate emails, which this function prepares to do.
    #
    # Due to heavy redundancy, the other "compose" functions will less thoroughly commented upon.

    # This names the first email after the folder and appends the word "FILM" for differentiation.

    film_subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1] + ' FILM'

    # This sets the body text, which is simpler in this case than in the heat transfer email.

    film_body_list = ['This order requires film.']

    # This block attaches the count and the film art pages.

    film_attachment_list = [csv_path]

    for file in os.listdir(folder_path):
        if 'Actual' in str(file):
            film_attachment_list.append(os.path.join(folder_path, file))

    # This calls the mail_maker function to cobble together an email from these elements.

    mail_maker(film_subject, film_body_list, film_attachment_list)

    # This names the second email after the folder and appends the words "HEAT TRANSFER" for differentiation.

    trans_subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1] + ' HEAT TRANSFER'

    # This sets the body text depending upon the specific requirements of the order.

    trans_body_list = ['This order also requires ']

    if needs_heat_transfers and needs_player_numbers:
        trans_body_list.append('heat transfers and player numbers.')
    elif needs_heat_transfers and has_player_numbers:
        trans_body_list.append('heat transfers. Its player numbers are covered by stock.')
    elif needs_heat_transfers:
        trans_body_list.append('heat transfers. No player numbers.')
    else:
        trans_body_list.append('player numbers. Its heat transfers are covered by stock.')

    if has_embroidery:
        trans_body_list.append('\n\nEmbroidered Items')

    # This block attaches the count sheet, the MLOrder file, and some combination of art pages and/or number sheets.

    trans_attachment_list = [csv_path]

    for file in os.listdir(folder_path):
        if 'MLOrder' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if 'ORDER Art Page' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '3N' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '4N' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '6N' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '8N' in str(file):
            trans_attachment_list.append(os.path.join(folder_path, file))

    # This calls the mail_maker function to cobble together an email from these elements.

    mail_maker(trans_subject, trans_body_list, trans_attachment_list)

    print('Emails generated: Combo')
    return


def compose_film(has_heat_transfers, has_player_numbers, has_embroidery, folder_path, csv_path):
    subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1]
    # This function is used for orders that need film elements but no heat transfer elements.
    # For more documentation, please see the comments on the "compose_combo" function.

    body_list = ['This order requires film. ']

    if has_heat_transfers and has_player_numbers:
        body_list.append('Its heat transfers and player numbers are covered by stock.')
    elif has_heat_transfers:
        body_list.append('Its heat transfers are covered by stock. No player numbers.')
    else:
        body_list.append('No heat transfers. No player numbers.')

    if has_embroidery:
        body_list.append('\n\nEmbroidered Items')

    attachment_list = [csv_path]

    for file in os.listdir(folder_path):
        if 'Actual' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '3N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '4N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '6N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '8N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    mail_maker(subject, body_list, attachment_list)

    print('Email generated: Film')
    return


def compose_transfer(needs_heat_transfers, has_player_numbers, needs_player_numbers, has_embroidery, folder_path,
                     csv_path):
    # This function is used for orders that need heat transfer elements but no film elements.
    # For more documentation, please see the comments on the "compose_combo" function.

    subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1]

    body_list = [' No film.']

    if needs_heat_transfers and needs_player_numbers:
        body_list.insert(0, 'This order requires heat transfers and player numbers.')
    elif needs_heat_transfers and has_player_numbers and not needs_player_numbers:
        body_list.insert(0, 'This order requires heat transfers. Its player numbers are covered by stock.')
    elif needs_heat_transfers and not has_player_numbers:
        body_list.insert(0, 'This order requires heat transfers. No player numbers.')
    elif needs_player_numbers and not needs_heat_transfers:
        body_list.insert(0, 'This order requires player numbers. Its heat transfers are covered by stock.')
    else:
        body_list.insert(0, 'This order\'s heat transfers and player numbers are covered by stock.')

    if has_embroidery:
        body_list.append('\n\nEmbroidered Items')

    attachment_list = [csv_path]

    for file in os.listdir(folder_path):
        if 'MLOrder' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if 'ORDER Art Page' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '3N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '4N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '6N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    for file in os.listdir(folder_path):
        if '8N' in str(file):
            attachment_list.append(os.path.join(folder_path, file))

    mail_maker(subject, body_list, attachment_list)

    print('Email generated: Heat Transfer')
    return


def compose_embroidery(folder_path, csv_path):
    # This function is used for orders that need exclusively embroidered items.
    # For more documentation, please see the comments on the "compose_combo" function.

    subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1]

    body_list = 'This order is all Embroidered Items. No heat transfers. No player numbers. No film.'

    attachment_list = [csv_path]

    mail_maker(subject, body_list, attachment_list)

    print('Email generated: Embroidery')
    return


def compose_stock(has_embroidery, folder_path, csv_path):
    # This function is used for orders that need nothing but do have heat transfers which are already in stock.
    # For more documentation, please see the comments on the "compose_combo" function.

    subject = folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1]

    body_list = ['This order is covered by stock. No player numbers. No film.']

    if has_embroidery:
        body_list.append(' \nEmbroidered Items')

    attachment_list = [csv_path]

    mail_maker(subject, body_list, attachment_list)

    print('Email generated: Stock')
    return


def mail_maker(subject, body_list, attachment_list):
    # This function takes the prepared email components from any one of the "compose" functions and creates an email.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)
    recipient_list = str(config['Outbot Options']['recipient_list']).split(',')
    auto_send = config.getboolean('Outbot Options', 'auto_send')

    outlook = win32.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    mail_item.Subject = subject

    mail_item.To = '; '.join(recipient_list)
    mail_item.Body = ''.join(body_list)
    mail_item.BodyFormat = 2

    for item in attachment_list:
        mail_item.Attachments.Add(item)

    # Depending on the configuration, the email is either sent or saved as a draft.

    if auto_send:
        mail_item.Send()
    else:
        mail_item.Save()
