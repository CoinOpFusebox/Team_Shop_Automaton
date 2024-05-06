# ---------------------------------------------------------------------------------------------------------------------#
# wts_order_processor.py
#
# This module runs all the other modules, processing a WTS order from inbox to outbox.
#
# ---------------------------------------------------------------------------------------------------------------------#

import atexit
import os
import re
import shutil
import win32com.client

from configparser import ConfigParser
from datetime import datetime
from pathlib import Path
from subprocess import run, DEVNULL
from time import sleep

import art_page_combobulator
import folder_molder
import player_number_cruncher
import ml_order_former
import log_borg
import wts_outbot


def main():
    # This block ensures that Illustrator and Excel are terminated when this program finishes running.
    # Without this, things rapidly descend into chaos.

    atexit.register(killustrator)
    atexit.register(excelcute)

    # This block opens the configuration file and retrieves the required folder/file paths.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)

    working_folder_path = config['Folder Paths']['working_folder_path']
    failed_folder_path = config['Folder Paths']['failed_folder_path']
    postponed_folder_path = config['Folder Paths']['postponed_folder_path']
    finished_folder_path = config['Folder Paths']['finished_folder_path']

    folder_path = ''
    number_order_list = []
    hta_count = 0

    # If there are folders in the Working Folder, this gets the path for the first one.
    # Otherwise, it calls the folder_molder to make a folder for the oldest email in the Postorders email folder.
    # If there are no postorder emails, the program ends.

    while not folder_path:
        if not len(os.listdir(working_folder_path)):
            folder_path = folder_molder.main()

            if folder_path == 'No emails!':
                return False
            elif folder_path == 'Blank count!':
                folder_path = None
                continue
        else:
            for dirpath, dirnames, filenames in os.walk(working_folder_path):
                if not dirnames:
                    folder_path = dirpath
                    break

    print(folder_path.split(os.path.sep)[-2] + ' ' + folder_path.split(os.path.sep)[-1])

    # This string holds the location of the Number Order List, which may or may not ever exist.

    order_list_path = ''.join((folder_path, r'\Number Order List.txt'))

    # These variables track which, if any, of the order's documents have already been generated.
    # If there are multiple files in the folder, in_medias_resonator is called to check if anything can be skipped.

    skip_heat_transfers = False
    skip_film = False
    skip_rdx = False
    skip_numbers = False
    skip_ml_order = False

    if len(os.listdir(folder_path)) > 1:
        skip_string = in_medias_resonator(folder_path)

        if skip_string:
            if 'h' in skip_string:
                skip_heat_transfers = True

            if 'f' in skip_string:
                skip_film = True

            if 'r' in skip_string:
                skip_rdx = True

            if 'n' in skip_string:
                skip_numbers = True

            if 'm' in skip_string:
                skip_ml_order = True

    # art_sheet_problem holds the status of an unsuccessful art page combobulation attempt, allowing the folder to be
    # moved appropriately.

    art_sheet_problem = False

    # This block calls the art_page_combobulator to create any required heat transfer or film art pages.

    if not skip_heat_transfers and not skip_film and not skip_rdx:
        try:
            art_sheet_problem = art_page_combobulator.combobulate(folder_path)
        except:
            fail(folder_path, working_folder_path, failed_folder_path)
            return True
    elif not skip_film and not skip_rdx:
        print("Heat transfer page already created!")
        try:
            art_sheet_problem = art_page_combobulator.combobulate(folder_path, skip_heat_transfers=True)
        except:
            fail(folder_path, working_folder_path, failed_folder_path)
            return True
    elif not skip_rdx:
        print("Heat transfer and/or film page(s) already created!")
        try:
            art_sheet_problem = art_page_combobulator.combobulate(folder_path, skip_heat_transfers=True, skip_film=True)
        except:
            fail(folder_path, working_folder_path, failed_folder_path)
            return True
    else:
        print("All art pages already created!")

    # If the combobulator was unsuccessful, this block moves the folder and terminates the program.

    if art_sheet_problem == 'Early':
        postpone(folder_path, working_folder_path, postponed_folder_path)
        return True
    elif art_sheet_problem:
        fail(folder_path, working_folder_path, failed_folder_path)
        return True

    # This block calls the player_number_cruncher to create any required player number sheets.

    if skip_numbers is False:
        try:
            number_order_list = player_number_cruncher.main(folder_path)
        except:
            fail(folder_path, working_folder_path, failed_folder_path)
            return True
    else:
        try:
            with open(order_list_path, 'r') as list_txt:
                txt_input_raw = list_txt.read()
                txt_input_string = ''.join(txt_input_raw)
                number_order_list = txt_input_string.split(',')
        except FileNotFoundError:
            try:
                number_order_list = player_number_cruncher.main(folder_path)
            except:
                fail(folder_path, working_folder_path, failed_folder_path)
                return True

    # This block calls the ml_order_former to create the order sheet for heat transfers and/or player numbers.

    if skip_ml_order is False:
        try:
            hta_count = ml_order_former.order_heat_transfers(folder_path)
        except:
            fail(folder_path, working_folder_path, failed_folder_path)
            return True

        if number_order_list:
            try:
                ml_order_former.order_player_numbers(folder_path, number_order_list, hta_count)
            except:
                fail(folder_path, working_folder_path, failed_folder_path)
                return True

        else:
            print('No player numbers to order!')
    else:
        print('MLOrder already created!')

    # This block calls the log_borg to log the order information in the Serigraphy database.

    try:
        log_borg.main(folder_path)
    except:
        fail(folder_path, working_folder_path, failed_folder_path)
        return True

    # This block calls the wts_outbot to generate emails for the order.

    try:
        wts_outbot.main(folder_path)
    except:
        fail(folder_path, working_folder_path, failed_folder_path)
        return True

    # This block deletes the Number Order List text file and moves the folder into 20XX Orders.

    sleep(10)

    try:
        os.remove(order_list_path)
    except FileNotFoundError:
        pass

    put_away(folder_path, finished_folder_path)

    print('Order processed successfully!\n')

    return True


def put_away(working_path, finished_folder_path):
    # This function moves a successful order's files into long-term storage.

    # working_path is the folder within which the files have been stored thus far.
    # finished_folder_path is the base destination folder.

    # destination_path is the complete and order-specific destination folder.
    # art_pgs_path is where most of the files go.
    # film_path is where the film files go.

    store_id = str(working_path.split(os.path.sep)[-1])
    destination_path = ''.join((finished_folder_path, '\\', str(datetime.now().year), ' Orders\\', store_id))
    # destination_path = ''.join((finished_folder_path, '\\2023 Orders\\', store_id))
    art_pgs_path = ''.join((destination_path, '\\Art Pgs'))

    film_path = ''

    while not film_path:
        for dirpath, dirnames, filenames in os.walk(destination_path):
            if 'Film' in dirpath:
                film_path = dirpath
                break

    while os.listdir(working_path):
        file = os.listdir(working_path)[0]

        old_path = ''.join((working_path, '\\', file))

        if 'Film' in file:
            new_path = ''.join((film_path, '\\', file))
        else:
            new_path = ''.join((art_pgs_path, '\\', file))

        shutil.move(old_path, new_path)

    if not os.listdir(working_path):
        os.rmdir(working_path)

    if not os.listdir(Path(working_path).parent):
        os.rmdir(Path(working_path).parent)

    print('Order moved successfully!')


def fail(folder_path, working_folder_path, failed_folder_path):
    # This function is called when something breaks.
    # It moves the folder into Failed Folders for review.

    killustrator_junior()
    excelcute()

    old_team_path = Path(folder_path).parent
    new_team_path = str(old_team_path).replace(working_folder_path, failed_folder_path)

    sleep(10)

    os.makedirs(new_team_path, exist_ok=True)
    shutil.move(str(folder_path), str(new_team_path))
    if not len(os.listdir(old_team_path)):
        os.rmdir(old_team_path)

    print('Order failed!\n')


def postpone(folder_path, working_folder_path, postponed_folder_path):
    # This function is called when film names/numbers are needed and not found.
    # It moves the folder into Postponed Folders.

    killustrator_junior()
    excelcute()

    old_team_path = Path(folder_path).parent
    new_team_path = str(old_team_path).replace(working_folder_path, postponed_folder_path)

    sleep(10)

    os.makedirs(new_team_path, exist_ok=True)
    shutil.move(str(folder_path), str(new_team_path))
    if not len(os.listdir(old_team_path)):
        os.rmdir(old_team_path)

    print('Order postponed!\n')


def in_medias_resonator(folder_path):
    # This function scans the folder for previously-completed documents to avoid needless repetition of work.

    skip_list = []
    number_order_sheet = False
    ml_order_sheet = False
    number_sheet = False

    for file in os.listdir(folder_path):
        if 'Order Art Page.' in file:
            skip_list.append('h')
            continue
        elif 'Film Art Page.' in file:
            skip_list.append('f')
            continue
        elif 'RDX Art Page.' in file:
            skip_list.append('r')
            continue
        elif 'Number Order List' in file:
            number_order_sheet = True
            continue
        elif 'MLOrder' in file and 'Part' not in file:
            ml_order_sheet = True
            continue
        elif number_sheet is False and re.search(r'\d[N]', file):
            number_sheet = True
            continue

    if number_sheet and number_order_sheet:
        skip_list.append('n')
    elif ml_order_sheet and not number_order_sheet:
        skip_list.append('nm')

    return ''.join(skip_list)


def killustrator():
    # This function closes Adobe Illustrator without asking questions.

    run('taskkill /im illustrator.exe /t /f', stderr=DEVNULL, stdout=DEVNULL)


def killustrator_junior():
    # This function closes all documents open in Adobe Illustrator, to provide a clean slate.

    illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
    illustrator.UserInteractionLevel = -1

    for item in illustrator.Documents:
        item.Close(2)


def excelcute():
    # This function closes Microsoft Excel, hard.

    run('taskkill /im excel.exe /t /f', stderr=DEVNULL, stdout=DEVNULL)


main()
