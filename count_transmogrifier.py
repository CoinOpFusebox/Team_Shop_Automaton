# ---------------------------------------------------------------------------------------------------------------------#
# count_transmogrifier.py
#
# This module takes a .CSV count file and returns a list of either heat transfers or film and their quantities.
#
# NOTE: This works using modified number templates that are missing the placeholder number block, which I should have
# gotten rid of years ago anyway.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import re
import win32com.client


def count_heat_transfers(csv_path):
    # csv_path contains the path for the count sheet.

    # outbound_list is a list that will contain a tuple for each HTA and its quantity.
    # If it remains empty, the order has no heat transfers.

    # inv_tuple is the most recently added HTA on the list.

    outbound_list = []
    second_chance_list = []
    inv_tuple = ()

    # This opens the count file.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # before_string and after_string contain the text that abuts the desired lines.

        before_string = '\'Heat Transfer\''
        after_string = '\'Total\''

        # These loops check for the before_string. When it's found, the loop ends.

        for line in reader:
            if before_string in str(line):
                break

        # This block checks for the after_string.
        # If it's not found, it converts the line into a tuple and adds it to the outbound_list.
        # When it's found, the loop ends.
        #
        # Along the way, it does some filtering to try to avoid bad entries. More filtering probably to come as new and
        # novel typos make themselves known!

        for line in reader:
            if after_string in str(line):
                break

            # This block is built to handle the occasional 'missing name' HTAs that have been showing up in counts
            # (for example, "HTA0081" instead of "GuggenheimCuratorsHTA0081").
            # If there's already a correctly-named HTA in the outbound list, its name will be borrowed and used.
            # Otherwise, the item is placed in a separate list to try again later.
            #
            # It now also accounts for accidental letters on the right side of "HTA" ("GuggenheimCuratorsHTAv0081",
            # for example), which have also started appearing.

            if 'HTA' in str(line[0]):
                hta_name = str(line[0])
                hta_number = int(line[1])

                if not hta_name.split('HTA')[1].isnumeric():
                    fixed_number = str(re.sub(r"\D+", "", hta_name.split('HTA')[1]))
                    hta_name = ''.join((hta_name.split('HTA')[0], 'HTA', fixed_number))

                if inv_tuple:
                    if inv_tuple[0] == str(rectifier(hta_name)):
                        outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number)
                        continue
                if len(hta_name.split('HTA')[0]) < 3:
                    if len(outbound_list):
                        borrowed_name = outbound_list[0][0].split('HTA')[0] + hta_name
                        if inv_tuple:
                            if inv_tuple[0] == str(rectifier(borrowed_name)):
                                outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number)
                                continue
                        inv_tuple = (str(rectifier(borrowed_name)), hta_number)
                        outbound_list.append(inv_tuple)
                    else:
                        inv_tuple = (hta_name, hta_number)
                        second_chance_list.append(inv_tuple)
                        continue

                else:
                    inv_tuple = (str(rectifier(hta_name)), int(hta_number))
                    outbound_list.append(inv_tuple)

        if len(second_chance_list):
            for item in second_chance_list:
                if len(outbound_list):
                    borrowed_name = outbound_list[0][0].split('HTA')[0] + item[0]
                    inv_tuple = (str(rectifier(borrowed_name)), int(item[1]))
                    outbound_list.append(inv_tuple)

    return outbound_list


def count_film(csv_path):
    # csv_path contains the path for the count sheet.

    # outbound_list is a list that will contain a tuple for each HTA, its quantity, and its film type.
    # If it remains empty, the order has none of the selected type of item.

    # inv_tuple is the most recently added HTA on the list.

    outbound_list = []
    second_chance_list = []
    inv_tuple = ()

    # This opens the count file.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # before_string and after_string contain the text that abuts the desired lines.

        before_string = '\'Film\''
        after_string = '\'Total\''

        # These loops check for the before_string. When it's found, the loop ends.

        for line in reader:
            if before_string in str(line):
                break

        # This block checks for the end_string.
        # If it's not found, it converts the line into a tuple and adds it to the outbound_list.
        # When it's found, the loop ends.
        #
        # Along the way, it does some filtering to try to avoid bad entries. More filtering probably to come as new and
        # novel typos make themselves known!

        names_and_numbers = False

        # This checks for film name and number elements.
        # I've added filtering to stop orders with nothing but "NONE" or "N/A" from being wrongfully postponed.
        # This will fail if a player is named "Anne" or "Nona" when there is also a "None", but c'est la vie.

        for line in reader:
            if 'glove'.casefold() in str(line).casefold():
                continue

            try:
                if line[5]:
                    for character in line[5]:
                        if character.isnumeric():
                            names_and_numbers = True
                            break
                elif line[4]:
                    if 'none'.casefold() in line[4].casefold() or '/' in line[4]:
                        none_set = set('aeon')
                        for character in line[4]:
                            if character.isalpha() and character not in none_set:
                                names_and_numbers = True
                                break
                    else:
                        names_and_numbers = True
            except IndexError:
                pass

            if after_string in str(line):
                break

            if any(word in str(line).casefold() for word in ['helmet'.casefold(), 'skull'.casefold()]):
                type_string = 'Helmet'
            else:
                type_string = 'Film'

            if 'HTA' in str(line[0]):
                hta_name = str(line[0])
                hta_number = int(line[1])

                if not hta_name.split('HTA')[1].isnumeric():
                    fixed_number = str(re.sub(r"\D+", "", hta_name.split('HTA')[1]))
                    hta_name = ''.join((hta_name.split('HTA')[0], 'HTA', fixed_number))

                if inv_tuple:
                    if inv_tuple[0] == str(rectifier(hta_name)) and inv_tuple[2] == type_string:
                        outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                             outbound_list[-1][2])
                        continue
                if len(hta_name.split('HTA')[0]) < 3:
                    if len(outbound_list):
                        borrowed_name = outbound_list[0][0].split('HTA')[0] + hta_name
                        if inv_tuple:
                            if inv_tuple[0] == str(rectifier(borrowed_name)) and inv_tuple[2] == type_string:
                                outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                                     outbound_list[-1][2])
                                continue
                        inv_tuple = (str(rectifier(borrowed_name)), hta_number, type_string)
                        outbound_list.append(inv_tuple)
                    else:
                        inv_tuple = (hta_name, hta_number, type_string)
                        second_chance_list.append(inv_tuple)
                        continue
                else:
                    inv_tuple = (str(rectifier(hta_name)), hta_number, type_string)
                    outbound_list.append(inv_tuple)

        # If there are any entries in the second chance list, this tries them again.

        if len(second_chance_list):
            for item in second_chance_list:
                if len(outbound_list):
                    borrowed_name = outbound_list[0][0].split('HTA')[0] + item[0]
                    inv_tuple = (str(rectifier(borrowed_name)), int(item[1]), str(item[2]))
                    outbound_list.append(inv_tuple)

        # If there are names and numbers, this checks for an associated file in Sublimation Artwork 2.
        # If it finds one, it checks the size. If it's shorter than an art page box, it adds an entry to the
        # outbound list noting the presence of names and numbers as well as a path to the file. If it's taller,
        # placeholder entries may first be added to the list to ensure that it doesn't run off the page.

        if names_and_numbers:
            try:
                sub_art_path = ''.join(('R:\\Sublimation Artwork 2\\', csv_path.split(os.path.sep)[-3], '\\',
                                        csv_path.split(os.path.sep)[-2]))

                name_and_numbers_found = False

                for file in os.listdir(sub_art_path):
                    if 'HTA' not in file and '.eps' in file:
                        names_and_numbers_path = ''.join((sub_art_path, '\\', str(file)))
                        illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
                        illustrator.UserInteractionLevel = -1
                        illustrator.Open(names_and_numbers_path)
                        names_and_numbers_document = illustrator.ActiveDocument
                        illustrator.ExecuteMenuCommand('selectall')
                        illustrator.ExecuteMenuCommand('group')
                        if names_and_numbers_document.Selection[0].Height <= 765:
                            outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                            names_and_numbers_document.Close(2)
                            name_and_numbers_found = True
                            break
                        else:
                            hta_box_count = len(outbound_list)
                            name_and_number_box_count = (names_and_numbers_document.Selection[0].Height // 765) + 1
                            if (hta_box_count + name_and_number_box_count) // 10 > hta_box_count // 10 \
                                    and (hta_box_count + name_and_number_box_count) % 10 != 0:
                                spacer_count = (((hta_box_count // 10) + 1) * 10) - hta_box_count
                                while spacer_count:
                                    outbound_list.append((' ', 0, ' '))
                                    spacer_count -= 1
                                names_and_numbers_document.Close(2)
                                name_and_numbers_found = True
                                break
                            else:
                                outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                                names_and_numbers_document.Close(2)
                                name_and_numbers_found = True
                                break

                if not name_and_numbers_found:
                    return 'Names and numbers not found!'
            except FileNotFoundError:
                return 'Names and numbers not found!'

    return outbound_list


def count_milb_film(csv_path):
    # csv_path contains the path for the count sheet.

    # outbound_list is a list that will contain a tuple for each HTA, its quantity, and its film type.

    # inv_tuple is the most recently added HTA on the list.

    outbound_list = []
    second_chance_list = []
    inv_tuple = ()

    # This opens the count file.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # before_string and after_string contain the text that abuts the desired lines.

        before_string = '\'Heat Transfer\''
        after_string = '\'Total\''

        # These loops check for the before_string. When it's found, the loop ends.

        for line in reader:
            if before_string in str(line):
                break

        # This block checks for the end_string.
        # If it's not found, it converts the line into a tuple and adds it to the outbound_list.
        # When it's found, the loop ends.
        #
        # Along the way, it does some filtering to try to avoid bad entries. More filtering probably to come as new and
        # novel typos make themselves known!

        for line in reader:
            if after_string in str(line):
                break

            type_string = 'Film'

            if 'HTA' in str(line[0]):
                hta_name = str(line[0])
                hta_number = int(line[1])

                if not hta_name.split('HTA')[1].isnumeric():
                    fixed_number = str(re.sub(r"\D+", "", hta_name.split('HTA')[1]))
                    hta_name = ''.join((hta_name.split('HTA')[0], 'HTA', fixed_number))

                if inv_tuple:
                    if inv_tuple[0] == str(rectifier(hta_name)) and inv_tuple[2] == type_string:
                        outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                             outbound_list[-1][2])
                        continue
                if len(hta_name.split('HTA')[0]) < 3:
                    if len(outbound_list):
                        borrowed_name = outbound_list[0][0].split('HTA')[0] + hta_name
                        if inv_tuple:
                            if inv_tuple[0] == str(rectifier(borrowed_name)) and inv_tuple[2] == type_string:
                                outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                                     outbound_list[-1][2])
                                continue
                        inv_tuple = (str(rectifier(borrowed_name)), hta_number, type_string)
                        outbound_list.append(inv_tuple)
                    else:
                        inv_tuple = (hta_name, hta_number, type_string)
                        second_chance_list.append(inv_tuple)
                        continue
                else:
                    inv_tuple = (str(rectifier(hta_name)), hta_number, type_string)
                    outbound_list.append(inv_tuple)

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # before_string and after_string contain the text that abuts the desired lines.

        before_string = '\'Film\''
        after_string = '\'Total\''

        # These loops check for the before_string. When it's found, the loop ends.

        for line in reader:
            if before_string in str(line):
                break

        # This block checks for the end_string.
        # If it's not found, it converts the line into a tuple and adds it to the outbound_list.
        # When it's found, the loop ends.
        #
        # Along the way, it does some filtering to try to avoid bad entries. More filtering probably to come as new and
        # novel typos make themselves known!

        names_and_numbers = False

        # This checks for film name and number elements.
        # I've added filtering to stop orders with nothing but "NONE" or "N/A" from being wrongfully postponed.
        # This will fail if a player is named "Anne" or "Nona" when there is also a "None", but c'est la vie.

        for line in reader:
            if 'glove'.casefold() in str(line).casefold():
                continue

            try:
                if line[5]:
                    for character in line[5]:
                        if character.isnumeric():
                            names_and_numbers = True
                            break
                elif line[4]:
                    if 'none'.casefold() in line[4].casefold() or '/' in line[4]:
                        none_set = set('aeon')
                        for character in line[4]:
                            if character.isalpha() and character not in none_set:
                                names_and_numbers = True
                                break
                    else:
                        names_and_numbers = True
            except IndexError:
                pass

            if after_string in str(line):
                break

            if any(word in str(line).casefold() for word in ['helmet'.casefold(), 'skull'.casefold()]):
                type_string = 'Helmet'
            else:
                type_string = 'Film'

            if 'HTA' in str(line[0]):
                hta_name = str(line[0])
                hta_number = int(line[1])

                if not hta_name.split('HTA')[1].isnumeric():
                    fixed_number = str(re.sub(r"\D+", "", hta_name.split('HTA')[1]))
                    hta_name = ''.join((hta_name.split('HTA')[0], 'HTA', fixed_number))

                if inv_tuple:
                    if inv_tuple[0] == str(rectifier(hta_name)) and inv_tuple[2] == type_string:
                        outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                             outbound_list[-1][2])
                        continue
                if len(hta_name.split('HTA')[0]) < 3:
                    if len(outbound_list):
                        borrowed_name = outbound_list[0][0].split('HTA')[0] + hta_name
                        if inv_tuple:
                            if inv_tuple[0] == str(rectifier(borrowed_name)) and inv_tuple[2] == type_string:
                                outbound_list[-1] = (outbound_list[-1][0], outbound_list[-1][1] + hta_number,
                                                     outbound_list[-1][2])
                                continue
                        inv_tuple = (str(rectifier(borrowed_name)), hta_number, type_string)
                        outbound_list.append(inv_tuple)
                    else:
                        inv_tuple = (hta_name, hta_number, type_string)
                        second_chance_list.append(inv_tuple)
                        continue
                else:
                    inv_tuple = (str(rectifier(hta_name)), hta_number, type_string)
                    outbound_list.append(inv_tuple)

        # If there are any entries in the second chance list, this tries them again.

        if len(second_chance_list):
            for item in second_chance_list:
                if len(outbound_list):
                    borrowed_name = outbound_list[0][0].split('HTA')[0] + item[0]
                    inv_tuple = (str(rectifier(borrowed_name)), int(item[1]), str(item[2]))
                    outbound_list.append(inv_tuple)

        # If there are names and numbers, this checks for an associated file in Sublimation Artwork 2.
        # If it finds one, it checks the size. If it's shorter than an art page box, it adds an entry to the
        # outbound list noting the presence of names and numbers as well as a path to the file. If it's taller,
        # placeholder entries may first be added to the list to ensure that it doesn't run off the page.

        outbound_list.sort()

        # This block combines matching HTAs.

        duplicate_list = []

        last_name = None
        last_quantity = None
        last_type = None

        for item in outbound_list:
            if not last_name:
                last_name = item[0]
                last_quantity = item[1]
                last_type = item[2]
                duplicate_list.append(item)
                continue
            if item[0] == last_name and item[2] == last_type:
                print(duplicate_list[-1][1])
                print(last_quantity)
                duplicate_list[-1] = (last_name, last_quantity + item[1], last_type)
                last_quantity = duplicate_list[-1][1]
                print(last_quantity)
            else:
                duplicate_list.append(item)
                last_name = item[0]
                last_quantity = item[1]
                last_type = item[2]

        outbound_list = duplicate_list

        if names_and_numbers:
            try:
                sub_art_path = ''.join(('R:\\Sublimation Artwork 2\\', csv_path.split(os.path.sep)[-3], '\\',
                                        csv_path.split(os.path.sep)[-2]))

                name_and_numbers_found = False

                for file in os.listdir(sub_art_path):
                    if 'HTA' not in file and '.eps' in file:
                        names_and_numbers_path = ''.join((sub_art_path, '\\', str(file)))
                        illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
                        illustrator.UserInteractionLevel = -1
                        illustrator.Open(names_and_numbers_path)
                        names_and_numbers_document = illustrator.ActiveDocument
                        illustrator.ExecuteMenuCommand('selectall')
                        illustrator.ExecuteMenuCommand('group')
                        if names_and_numbers_document.Selection[0].Height <= 765:
                            outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                            names_and_numbers_document.Close(2)
                            name_and_numbers_found = True
                            break
                        else:
                            hta_box_count = len(outbound_list)
                            name_and_number_box_count = (names_and_numbers_document.Selection[0].Height // 765) + 1
                            if (hta_box_count + name_and_number_box_count) // 10 > hta_box_count // 10 \
                                    and (hta_box_count + name_and_number_box_count) % 10 != 0:
                                spacer_count = (((hta_box_count // 10) + 1) * 10) - hta_box_count
                                while spacer_count:
                                    outbound_list.append((' ', 0, ' '))
                                    spacer_count -= 1
                                names_and_numbers_document.Close(2)
                                name_and_numbers_found = True
                                break
                            else:
                                outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                                names_and_numbers_document.Close(2)
                                name_and_numbers_found = True
                                break

                if not name_and_numbers_found:
                    return 'Names and numbers not found!'
            except FileNotFoundError:
                return 'Names and numbers not found!'

    return outbound_list


def rectifier(whole_hta_string):
    # This function makes sure that the HTA number has the proper number of digits, which is four.
    # It does this by adding or removing leading zeroes.
    # Now with 20% more not crashing on film pieces with no HTA!
    #
    # The quadifier is now the rectifier, which also corrects common issues and typos.

    whole_hta_list = re.split("HTA", whole_hta_string)

    hta_digits = whole_hta_list[1]

    if len(hta_digits) > 4:
        cut_digits = len(hta_digits) - 4
        hta_digits = hta_digits[cut_digits:]
    elif len(hta_digits) < 4:
        add_digits = 4 - len(hta_digits)
        while add_digits > 0:
            hta_digits = '0' + hta_digits
            add_digits = add_digits - 1

    whole_hta_string = whole_hta_list[0] + 'HTA' + hta_digits

    if '\'' in whole_hta_string:
        whole_hta_string = whole_hta_string.replace('\'', '’')

    if 'Customer Style' in whole_hta_string:
        whole_hta_string = whole_hta_string.replace('Customer Style', '')

    if 'CustomerStyle' in whole_hta_string:
        whole_hta_string = whole_hta_string.replace('CustomerStyle', '')

    return whole_hta_string
