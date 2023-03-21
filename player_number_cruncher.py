# ---------------------------------------------------------------------------------------------------------------------#
# player_number_cruncher.py
#
# This module checks for and processes player numbers and their art sheets.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import pyclip
import pyodbc
import re
import wilson_colors
import win32com.client

from configparser import ConfigParser
from dotenv import load_dotenv
from pathlib import Path
from time import sleep


config_path = Path(__file__).parent.absolute().joinpath('config.ini')
config = ConfigParser()
config.read(config_path)

number_path = config['Folder Paths']['number_path']

# This dictionary matches abbreviated colors with their full names and their short names.

chromotome = {
    'Blk': 'Black',
    'Vin': 'Vintage White',
    'Wht': 'White',
    'Nav': 'Navy',
    'Sca': 'Scarlet',
    'Chr': 'Charcoal',
    'Slg': 'Silver',
    'Blg': 'Blue Grey',
    'Vic': 'Victory Blue',
    'Roy': 'Royal Blue',
    'Veg': 'Vegas Gold',
    'Lig': 'Light Gold',
    'Old': 'Old Gold',
    'Grl': 'Grello',
    'Bor': 'Burnt Orange',
    'Crd': 'Cardinal',
    'Dkb': 'Dark Brown',
    'Dkg': 'Dark Green',
    'Kel': 'Kelly Green',
    'Tno': 'TN Orange',
    'Tex': 'TX Orange',
    'Pur': 'Purple',
    'Mnb': 'Mineral Blue',
    'Mar': 'Maroon',
    'Lem': 'Lemon Yellow',
    'Hpk': 'Hot Pink',
    'Pop': 'Popping Pink',
    'Pnk': 'Pink',
    'Bkr': 'Brick Red',
    'Pan': 'Panther Blue',
    'Sgr': 'Speed Green',
    'Tel': 'Teal',
    'Mng': 'Mean Green',
    'Lim': 'Lime Green',
    'Hog': 'Hyper Orange',
    'Pyl': 'Punch Yellow',
    'Ldv': 'Loud Violet',
    'Ebl': 'Electric Blue',
    'Sth': 'Show-Through'
}


def main(folder_path):
    csv_path = ''

    # This block finds the .CSV count file in the folder, if there is one, then checks it for numbers.

    for file in os.listdir(folder_path):
        if file.endswith('.csv'):
            csv_path = os.path.join(folder_path, file)
            break

    # number_section is a list of the germane lines from the count.

    number_section = []

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # number_time indicates that the player numbers section has been reached.
        #
        # has_digits is True if numeric characters are found in the HT Number column.
        # This filters out orders whose garment designs feature numbers but the garments ordered do not.
        #
        # has_player is True if the word "Player" is found in the HT Description column
        # This filters out non-heat-transfer numbers, such as the ones found on ProFusion garments.

        number_time = False
        has_digits = False
        has_player = False

        for line in reader:
            if 'HT Adult CID' in line:
                number_time = True
            if number_time:
                number_section.append(line)
                if len(line) >= 5:
                    if re.search(r'\d', line[4]):
                        has_digits = True
                    if 'Player' in line[3]:
                        has_player = True

    # A section with four or fewer lines, no digits, or a description without "Player" is devoid of player numbers.

    if not (has_digits and has_player and len(number_section) >= 4):
        print('No player numbers!')
        return None

    player_number_list = []

    # Youth sizes will be used when kid_mode is active.

    kid_mode = False

    for line in number_section:
        if len(line) >= 5:
            if 'Player' in line[3] and re.search(r'\d', line[4]):
                crunch_tuple = crunch(line, kid_mode, player_number_list, folder_path)

                if crunch_tuple[0]:
                    player_number_list = crunch_tuple[0]

                player_number_list.append(crunch_tuple[1])

                if len(crunch_tuple) == 3:
                    player_number_list.append(crunch_tuple[2])

                continue

            elif 'Adult' in line[0]:
                kid_mode = False
                continue

            elif 'Youth' in line[0]:
                kid_mode = True
                continue
        else:
            continue

    # This sorts the list of PlayerNumbers objects in order of font/color, then by size, to ease later processing.

    player_number_list.sort(key=lambda numbers: ''.join((numbers.font, str(numbers.color_1), str(numbers.color_2),
                                                         str(numbers.color_3), numbers.size)))

    grand_digit_total = 0

    for item in player_number_list:
        grand_digit_total += item.total_digits

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)
        needs_total = True

        for line in reader:
            if '# Total' in str(line):
                needs_total = False
                break

    if needs_total:
        with open(csv_path, 'a', errors="ignore") as csv_count:
            total_row = ['# Total', str(grand_digit_total)]

            writer = csv.writer(csv_count)
            writer.writerow(total_row)

    order_list = craft(player_number_list, folder_path, csv_path)

    print('Player numbers processed!')
    return order_list


def crunch(line, kid_mode, crunched_list, folder_path):
    # This method turns a line in the number section into a full-fledged PlayerNumbers object.

    # sparta_notes hold information about the player number font, sizing, and colors.

    sparta_notes = line[3]

    # Sometimes extra information ends up in the Sparta Notes field. This removes it.

    if 'player'.casefold() in sparta_notes.casefold():
        sparta_notes = sparta_notes.split('player'.casefold())[-1]

    # This block gets the list of player numbers and counts the digits.

    back_number_list = []
    front_number_list = []

    for item in line[4].split(','):
        if item.isnumeric():
            back_number_list.append(item)

    back_total_digits = len(''.join(back_number_list))

    # This block gets the size(s).
    # The only sizing oddity currently accounted for is the larger front numbers found on some youth garments.

    front_size = 0

    if kid_mode:
        if '3' in sparta_notes and '4' in sparta_notes:
            back_size = '6N'
            front_size = '3N'
        elif '4' in sparta_notes:
            back_size = '6N'
            front_size = '4N'
        else:
            back_size = '6N'
    else:
        if '4' in sparta_notes:
            back_size = '8N'
            front_size = '4N'
        else:
            back_size = '8N'

    if front_size:
        for number in back_number_list:
            front_number_list.append(number)
            front_total_digits = len(''.join(front_number_list))

    # This block gets the font.
    # I've included the four normal fonts for now; other fonts will be added later (or dealt with manually).

    if any(word in sparta_notes.casefold() for word in ['full'.casefold(), 'fbk'.casefold()]):
        font = 'Fbk'
    elif any(word in sparta_notes.casefold() for word in ['fancy'.casefold(), 'fan'.casefold()]):
        font = 'Fan'
    elif any(word in sparta_notes.casefold() for word in ['vortex'.casefold(), 'vor'.casefold()]):
        font = 'Vor'
    elif 'nca'.casefold() in sparta_notes.casefold():
        font = 'Nca'
    elif any(word in sparta_notes.casefold() for word in ['arizona'.casefold(), 'ari'.casefold()]):
        font = 'Ari'
    elif 'custom'.casefold() in sparta_notes.casefold():
        if 'canes'.casefold() in folder_path.casefold():
            font = 'Ari'
        else:
            font = '???'
    else:
        font = '???'

    number_of_colors = chromatic_enumerator(sparta_notes)

    # This block gets the colors.

    color_string = sparta_notes.split('-C')[-1].split('A')[0]

    if number_of_colors == 1:
        color_1 = color_picker(color_string)
        color_2 = 0
        color_3 = 0

    elif number_of_colors == 2:
        color_split_list = color_string.split('/')

        color_1 = color_picker(color_split_list[0])
        color_2 = color_picker(color_split_list[1])
        color_3 = 0

    else:
        color_split_list = color_string.split('/')

        color_1 = color_picker(color_split_list[0])
        color_2 = color_picker(color_split_list[1])
        color_3 = color_picker(color_split_list[2])

    # This block makes sure that orders with number parameters I've yet to address (different front and back styles or
    # unknown colors/fonts, for example) fail gracefully.

    if '?' in ''.join((font, str(color_1), str(color_2), str(color_3))):
        raise ValueError('Unknown font or color!')

    # This block combines like numbers (for example, two stacks of 8N Fbk Wht).

    for item in crunched_list:
        if font == item.font and back_size == item.size and color_1 == item.color_1 and color_2 == item.color_2 and \
                color_3 == item.color_3:
            for number in item.number_list:
                back_number_list.append(number)

            back_total_digits = back_total_digits + item.total_digits

            crunched_list.remove(item)

    if front_size:
        for item in crunched_list:
            if font == item.font and front_size == item.size and color_1 == item.color_1 and color_2 == item.color_2 \
                    and color_3 == item.color_3:
                for number in item.number_list:
                    front_number_list.append(number)

                front_total_digits = front_total_digits + item.total_digits

                crunched_list.remove(item)

    back_number_list.sort(key=int)

    if front_size:
        front_number_list.sort(key=int)

    # This block opens the environment file and a connection to the database
    # For more info on interacting with the database, see the note at the top of count_comparitron.main().

    load_dotenv()

    connection_list = ['DSN=Heat Transfer Inventory;Database=Heat Transfer Inventory;UID=',
                       os.getenv('HTI_UID'),
                       ';PWD=',
                       os.getenv('HTI_PWD')]

    connection = pyodbc.connect(''.join(connection_list))
    cursor = connection.cursor()

    # This block gets the digit counts and checks the stock levels of the numbers.

    back_number_slurry = ''.join(back_number_list)
    if front_size:
        front_number_slurry = ''.join(front_number_list)

    back_digit_count_list = []
    front_digit_count_list = []
    digit = 0

    in_stock = True

    while digit <= 9:
        back_digit_count_list.append(back_number_slurry.count(str(digit)))
        digit += 1

    if front_size:
        digit = 0

        while digit <= 9:
            front_digit_count_list.append(front_number_slurry.count(str(digit)))
            digit += 1

    counter = 0

    if front_size:
        if number_of_colors == 1:
            back_filename = ' '.join((back_size, font, color_1))
            front_filename = ' '.join((front_size, font, color_1))

        elif number_of_colors == 2:
            back_filename = ' '.join((back_size, font, color_1, color_2))
            front_filename = ' '.join((front_size, font, color_1, color_2))

        else:
            back_filename = ' '.join((back_size, font, color_1, color_2, color_3))
            front_filename = ' '.join((front_size, font, color_1, color_2, color_3))

        back_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', back_filename).fetchone()
        front_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', front_filename).fetchone()

        if not back_row:
            back_filename = ''.join((back_filename, '.png'))
            back_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', back_filename).fetchone()

        if not front_row:
            front_filename = ''.join((front_filename, '.png'))
            front_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', front_filename).fetchone()

        while counter <= 9:
            on_hand_string = ''.join(('#', str(counter), ' Units On Hand'))

            try:
                if int(back_row.__getattribute__(on_hand_string)) < back_digit_count_list[counter]:
                    in_stock = False

                if int(front_row.__getattribute__(on_hand_string)) < front_digit_count_list[counter]:
                    in_stock = False
            except AttributeError:
                in_stock = False
                break

            if not in_stock:
                break

            counter += 1
    else:
        if number_of_colors == 1:
            back_filename = ' '.join((back_size, font, color_1))

        elif number_of_colors == 2:
            back_filename = ' '.join((back_size, font, color_1, color_2))

        else:
            back_filename = ' '.join((back_size, font, color_1, color_2, color_3))

        back_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', back_filename).fetchone()

        if not back_row:
            back_filename = ''.join((back_filename, '.png'))
            back_row = cursor.execute('SELECT * FROM Numbers WHERE Filename=?', back_filename).fetchone()

        while counter <= 9:
            on_hand_string = ''.join(('#', str(counter), ' Units On Hand'))

            try:
                if int(back_row.__getattribute__(on_hand_string)) < back_digit_count_list[counter]:
                    in_stock = False
                    break
            except AttributeError:
                in_stock = False
                break

            counter += 1

    if front_size:
        return (crunched_list,
                PlayerNumbers(back_number_list, back_total_digits, back_digit_count_list, back_size, font, color_1,
                              color_2, color_3, in_stock),
                PlayerNumbers(front_number_list, front_total_digits, front_digit_count_list, front_size, font, color_1,
                              color_2, color_3, in_stock))

    else:
        return (crunched_list, PlayerNumbers(back_number_list, back_total_digits, back_digit_count_list, back_size,
                                             font, color_1, color_2, color_3, in_stock))


def craft(player_number_list, folder_path, csv_path):
    # This method takes the player_numbers_list, makes player number art sheets, and updates the count file.

    # This block opens an Excel workbook within which data may be temporarily stored.

    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    workbook = excel.Workbooks.Add()
    top_holding_sheet = workbook.Worksheets.Add()
    top_holding_sheet.Name = 'Top Holding Sheet'
    bottom_holding_sheet = workbook.Worksheets.Add()
    bottom_holding_sheet.Name = 'Bottom Holding Sheet'
    working_sheet = workbook.Worksheets.Add()
    working_sheet.Name = 'Working Sheet'

    # max_y_counter holds the height of the tallest stack of numbers so that the section labels can be placed properly.
    # number_block_counter holds the number of number blocks already placed on the holding sheet.
    # all_in_stock is set to False if a font needs ordering.
    # order_list holds the data blocks for numbers that are not in stock so that they can be placed on the MLOrder.

    max_y_counter = 7
    number_block_counter = 0
    all_in_stock = True
    order_list = []

    for player_numbers in player_number_list:
        # This block gets the path for the template file.

        if player_numbers.color_3:
            color_count = '3-C'
        elif player_numbers.color_2:
            color_count = '2-C'
        else:
            color_count = '1-C'

        template_path = ''.join(
            (number_path, '\\', player_numbers.font.upper(), '\\', color_count, '\\', player_numbers.size, ' ',
             player_numbers.font, '.ai'))

        # This block selects a special path for two-color Full Block show-through numbers.
        # If I ever see another kind of show-through numbers I'll deal with it then.

        if player_numbers.color_1 == 'Sth':
            template_path = ''.join((number_path, '\\', player_numbers.font.upper(), '\\STH\\', player_numbers.size,
                                     ' ', player_numbers.font, '.ai'))

        # This block opens Illustrator and the template file.

        illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
        illustrator.UserInteractionLevel = -1

        sleep(10)

        illustrator.Open(template_path)
        number_sheet = illustrator.ActiveDocument

        # This block assembles and inserts the header.

        header = ''.join((folder_path.split(os.path.sep)[-2], ' ', folder_path.split(os.path.sep)[-1]))

        for item in number_sheet.TextFrames:
            if item.Contents == 'Fakeville Baseballers 19Q 0123456789':
                item.Contents = header
                break

        # This block assembles and inserts the info block.

        if player_numbers.font == 'Fbk':
            long_font = 'Full Block'
        elif player_numbers.font == 'Fan':
            long_font = 'Fancy Block'
        elif player_numbers.font == 'Vor':
            long_font = 'Vortex'
        elif player_numbers.font == 'Nca':
            long_font = 'NCAA'
        elif player_numbers.font == 'Ari':
            long_font = 'Arizona'
        else:
            long_font = '???'

        info_block = ''.join((long_font, '\n', player_numbers.size.replace('N', '"'), ' ', color_count, '\n',
                              chromotome[player_numbers.color_1]))

        if player_numbers.color_2:
            info_block = ''.join((info_block, '/', chromotome[player_numbers.color_2]))

        if player_numbers.color_3:
            info_block = ''.join((info_block, '/', chromotome[player_numbers.color_3]))

        for item in number_sheet.TextFrames:
            if 'Mint' in item.Contents:
                item.Contents = info_block
                break

        # This block sets the colors.

        for group in number_sheet.GroupItems:
            if group.Name == 'Color 1':
                wilson_color = getattr(wilson_colors, chromotome[player_numbers.color_1].replace(' ', '_').lower())
                for item in group.PathItems:
                    item.SetFillColor(wilson_color)
                    sleep(.2)
                for item in group.CompoundPathItems:
                    for sub_item in item.PathItems:
                        sub_item.SetFillColor(wilson_color)
                        sleep(.2)

            if group.Name == 'Color 2':
                wilson_color = getattr(wilson_colors, chromotome[player_numbers.color_2].replace(' ', '_').lower())
                for item in group.PathItems:
                    item.SetFillColor(wilson_color)
                    sleep(.2)
                for item in group.CompoundPathItems:
                    for sub_item in item.PathItems:
                        sub_item.SetFillColor(wilson_color)
                        sleep(.2)

            if group.Name == 'Color 3':
                wilson_color = getattr(wilson_colors, chromotome[player_numbers.color_3].replace(' ', '_').lower())
                for item in group.PathItems:
                    item.SetFillColor(wilson_color)
                    sleep(.2)
                for item in group.CompoundPathItems:
                    for sub_item in item.PathItems:
                        sub_item.SetFillColor(wilson_color)
                        sleep(.2)

        # This block inserts the number block into the number sheet.
        # Along the way, it stores the number data on the holding sheet for later insertion into the count sheet.

        # This first chunk creates the header.

        working_sheet.Cells(1, 1).Value = ' '.join((player_numbers.font, str(player_numbers.color_1),
                                                    str(player_numbers.color_2),
                                                    str(player_numbers.color_3))).replace('0', '').rstrip()
        working_sheet.Cells(3, 1).Value = player_numbers.size.replace('N', '"')
        working_sheet.Cells(3, 2).Value = 'Need'

        # y_counter keeps track of the height of the "actual player's numbers" section.
        # last_number is the number most recently added. Since the list has been sorted by crunch(), this allows for the
        # combination of same numbers.

        y_counter = 5
        last_number = -1

        for number in player_numbers.number_list:
            if number == last_number:
                working_sheet.Cells(y_counter - 1, 2).Value = working_sheet.Cells(y_counter - 1, 2).Value + 1
                continue

            working_sheet.Cells(y_counter, 1).Value = number
            working_sheet.Cells(y_counter, 2).Value = 1

            y_counter += 1

            last_number = number

        y_counter += 1

        if y_counter > max_y_counter:
            max_y_counter = y_counter

        # d_counter keeps track of the current digit being used.

        d_counter = 0

        while d_counter <= 9:
            working_sheet.Cells(y_counter + d_counter, 1).Value = d_counter
            working_sheet.Cells(y_counter + d_counter, 2).Value = player_numbers.digit_count_list[d_counter]

            d_counter += 1

        pyclip.copy('')

        if player_numbers.in_stock:
            while not pyclip.paste():
                sleep(1)
                working_sheet.Range(working_sheet.Cells(5, 1), working_sheet.Cells(y_counter - 2, 2)).\
                    HorizontalAlignment = -4117
                working_sheet.Range(working_sheet.Cells(5, 1), working_sheet.Cells(y_counter - 2, 2)).Copy()

        else:
            all_in_stock = False
            order_list.append(info_block)

            d_counter = 0

            while d_counter <= 9:
                if player_numbers.digit_count_list[d_counter] <= 6:
                    order_value = 10
                else:
                    order_value = (5 * round((player_numbers.digit_count_list[d_counter] * 1.5) / 5))

                working_sheet.Cells(max_y_counter + d_counter + 11, 1).Value = d_counter
                working_sheet.Cells(max_y_counter + d_counter + 11, 2).Value = order_value

                d_counter += 1

            while not pyclip.paste():
                sleep(1)
                working_sheet.Range(working_sheet.Cells(max_y_counter + 11, 1),
                                    working_sheet.Cells(max_y_counter + 20, 2)).HorizontalAlignment = -4117
                working_sheet.Range(working_sheet.Cells(max_y_counter + 11, 1),
                                    working_sheet.Cells(max_y_counter + 20, 2)).Copy()

        # This block pastes the number block (the actual player numbers, if in stock, or the ordering quantities if not)
        # into the number sheet and puts it in its proper place.

        for item in number_sheet.GroupItems:
            if item.Name == "Player Numbers Header":
                top_text = item
                break

        illustrator.ExecuteMenuCommand('paste')

        try:
            scale = 100 * (top_text.Width / number_sheet.Selection[0].Width)
        except TypeError:
            illustrator.ExecuteMenuCommand('paste')
            scale = 100 * (top_text.Width / number_sheet.Selection[0].Width)

        number_sheet.Selection[0].Resize(scale, scale)
        number_sheet.Selection[0].Left = top_text.Left
        number_sheet.Selection[0].Top = top_text.Top - 120

        # This saves and closes the number sheet.

        if player_numbers.color_3:
            save_path = ''.join((folder_path, '\\', folder_path.split(os.path.sep)[-2], ' ',
                                 folder_path.split(os.path.sep)[-1], ' ', player_numbers.size, ' ', player_numbers.font,
                                 ' ', player_numbers.color_1, ' ', player_numbers.color_2, ' ', player_numbers.color_3,
                                 '.ai'))
        elif player_numbers.color_2:
            save_path = ''.join((folder_path, '\\', folder_path.split(os.path.sep)[-2], ' ',
                                 folder_path.split(os.path.sep)[-1], ' ', player_numbers.size, ' ', player_numbers.font,
                                 ' ', player_numbers.color_1, ' ', player_numbers.color_2, '.ai'))
        else:
            save_path = ''.join((folder_path, '\\', folder_path.split(os.path.sep)[-2], ' ',
                                 folder_path.split(os.path.sep)[-1], ' ', player_numbers.size, ' ', player_numbers.font,
                                 ' ', player_numbers.color_1, '.ai'))

        ai_save_options = win32com.client.Dispatch('Illustrator.IllustratorSaveOptions')
        ai_save_options.Compatibility = 15
        number_sheet.SaveAs(save_path, ai_save_options)

        save_path = save_path.replace('.ai', '.pdf')
        pdf_save_options = win32com.client.Dispatch('Illustrator.PDFSaveOptions')
        number_sheet.SaveAs(save_path, pdf_save_options)

        number_sheet.Close(2)

        # This block moves the contents of the working sheet onto the holding sheets to make way for the next iteration.

        working_sheet.Range(working_sheet.Cells(1, 1), working_sheet.Cells(y_counter - 2, 2)).Cut(
            top_holding_sheet.Cells(1, (4 * number_block_counter) + 2))

        working_sheet.UsedRange.Cut(bottom_holding_sheet.Cells(1, (4 * number_block_counter) + 2))

        number_block_counter += 1

    # This block adds labels to the holding sheet.

    top_holding_sheet.Cells(5, 1).Value = "NUMBERS"
    top_holding_sheet.Cells(max_y_counter, 1).Value = "COUNT"

    if not all_in_stock:
        top_holding_sheet.Cells(max_y_counter + 11, 1).Value = "ORDER"

    # This block combines the data from the holding sheets, moves it onto the count sheet, and saves it.

    bottom_holding_sheet.UsedRange.Cut(top_holding_sheet.Cells(max_y_counter, 2))

    count_sheet = excel.Workbooks.Open(csv_path)

    top_holding_sheet.UsedRange.Cut(count_sheet.Worksheets[1].Cells(2, 8))

    excel.DisplayAlerts = False
    count_sheet.SaveAs(csv_path)
    excel.Quit()

    # If there are player numbers that need ordering, this creates a text file containing the order list.
    # This will make it possible to add player numbers to the MLOrder without re-running the cruncher.

    if order_list:
        with open(folder_path + r'\Number Order List.txt', 'w') as number_order_list:
            print(','.join(order_list).rstrip(), file=number_order_list)

    return order_list


def chromatic_enumerator(sparta_notes):
    count_string = sparta_notes.split('A')[0]

    if '1-C' in count_string:
        c_color_count = 1
    elif '2-C' in count_string:
        c_color_count = 2
    elif '3-C' in count_string:
        c_color_count = 3
    elif '1' in count_string:
        c_color_count = 1
    elif '2' in count_string:
        c_color_count = 2
    else:
        c_color_count = 3

    s_color_count = 1

    for character in count_string:
        if character == '/':
            s_color_count += 1

    if c_color_count == s_color_count:
        return c_color_count
    elif s_color_count <= 3:
        return s_color_count
    else:
        return c_color_count


def color_picker(color_string):
    # This function takes a string and looks for a color in it.
    # These are sorted in rough order of popularity to avoid running as many of these statements as possible.
    # I am entirely certain there's a more efficient way to do this.

    if 'black'.casefold() in color_string.casefold():
        return 'Blk'
    elif 'vintage'.casefold() in color_string.casefold():
        return 'Vin'
    elif 'white'.casefold() in color_string.casefold():
        return 'Wht'
    elif 'navy'.casefold() in color_string.casefold():
        return 'Nav'
    elif 'scarlet'.casefold() in color_string.casefold():
        return 'Sca'
    elif 'charcoal'.casefold() in color_string.casefold():
        return 'Chr'
    elif 'silver'.casefold() in color_string.casefold():
        return 'Slg'
    elif 'grey'.casefold() in color_string.casefold() or 'gray'.casefold() in color_string.casefold():
        return 'Blg'
    elif 'vic'.casefold() in color_string.casefold():
        return 'Vic'
    elif 'royal'.casefold() in color_string.casefold():
        return 'Roy'
    elif 'vegas'.casefold() in color_string.casefold():
        return 'Veg'
    elif 'gold'.casefold() in color_string.casefold():
        if 'light'.casefold() in color_string.casefold() or 'lt'.casefold() in color_string.casefold():
            return 'Lig'
        else:
            return 'Old'
    elif 'grello'.casefold() in color_string.casefold():
        return 'Grl'
    elif 'burnt'.casefold() in color_string.casefold() or 'bt'.casefold() in color_string.casefold():
        return 'Bor'
    elif 'cardinal'.casefold() in color_string.casefold():
        return 'Crd'
    elif 'brown'.casefold() in color_string.casefold():
        return 'Dkb'
    elif 'dark'.casefold() in color_string.casefold() or 'dk'.casefold() in color_string.casefold():
        return 'Dkg'
    elif 'kelly'.casefold() in color_string.casefold():
        return 'Kel'
    elif 'tennessee'.casefold() in color_string.casefold() or 'tn'.casefold() in color_string.casefold():
        return 'Tno'
    elif 'texas'.casefold() in color_string.casefold() or 'tx'.casefold() in color_string.casefold():
        return 'Tex'
    elif 'purple'.casefold() in color_string.casefold():
        return 'Pur'
    elif 'mineral'.casefold() in color_string.casefold():
        return 'Mnb'
    elif 'maroon'.casefold() in color_string.casefold():
        return 'Mar'
    elif 'lemon'.casefold() in color_string.casefold():
        return 'Lem'
    elif 'hot'.casefold() in color_string.casefold():
        return 'Hpk'
    elif 'pop'.casefold() in color_string.casefold():
        return 'Pop'
    elif 'pink'.casefold() in color_string.casefold():
        return 'Pnk'
    elif 'brick'.casefold() in color_string.casefold():
        return 'Bkr'
    elif 'panther'.casefold() in color_string.casefold():
        return 'Pan'
    elif 'speed'.casefold() in color_string.casefold():
        return 'Sgr'
    elif 'teal'.casefold() in color_string.casefold():
        return 'Tel'
    elif 'mean'.casefold() in color_string.casefold():
        return 'Mng'
    elif 'lime'.casefold() in color_string.casefold():
        return 'Lim'
    elif 'hyper'.casefold() in color_string.casefold():
        return 'Hog'
    elif 'punch'.casefold() in color_string.casefold():
        return 'Pyl'
    elif 'loud'.casefold() in color_string.casefold():
        return 'Ldv'
    elif 'elec'.casefold() in color_string.casefold():
        return 'Ebl'
    elif 'show'.casefold() in color_string.casefold():
        return 'Sho'
    else:
        return '???'


class PlayerNumbers:
    """A set of player numbers and their attributes."""

    def __init__(self, number_list, total_digits, digit_count_list, size, font, color_1, color_2, color_3, in_stock):
        self.number_list = number_list
        self.total_digits = total_digits
        self.digit_count_list = digit_count_list
        self.size = size
        self.font = font
        self.color_1 = color_1
        self.color_2 = color_2
        self.color_3 = color_3
        self.in_stock = in_stock
