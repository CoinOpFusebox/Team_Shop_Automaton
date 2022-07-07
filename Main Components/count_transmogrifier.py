# ---------------------------------------------------------------------------------------------------------------------#
# count_transmogrifier.py
#
# This module takes a .CSV count file and returns a list of either heat transfers or film and their quantities.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import re
import win32com.client

def main(csv_path, count_film):
    # csv_path contains the path for the count sheet.
    # If count_film is True, film HTAs will be counted. An additional entry will be added to each tuple to denote type.

    # outbound_list is a list that will contain a tuple for each HTA, count, and sometimes film type.
    # If it remains empty, the order has none of the selected type of item.

    outbound_list = []
    second_chance_list = []

    # This opens the count file.

    with open(csv_path, 'r', errors="ignore") as csv_count:
        reader = csv.reader(csv_count)

        # before_string and after_string contain the text that abuts the desired lines.

        if count_film is False:
            before_string = '\'Heat Transfer\''
        else:
            before_string = '\'Film\''

        after_string = '\'Total\''

        # This loop checks for the before_string. When it's found, the loop ends.

        for line in reader:
            if before_string in str(line):
                break

        # This block checks for the end_string.
        # If it's not found, it converts the line into a tuple and adds it to the outbound_list.
        # When it's found, the loop ends.
        #
        # Along the way, it does some filtering to try to avoid bad entries. More filtering probably to come as new and
        # novel typos make themselves known!

        if count_film is False:
            for line in reader:
                if after_string in str(line):
                    break

                # This block and its counterpart in the film loop are built to handle the occasional 'missing name' HTAs
                # that have been showing up in counts (for example, "HTA081" instead of "GuggenheimCuratorsHTA0081").
                # If there's already a correctly-named HTA in the outbound list, its name will be borrowed and used.
                # Otherwise, the item is placed in a separate list to try again later.

                if 'HTA' in str(line[0]):
                    if line[0].split('HTA')[0] == '':
                        if len(outbound_list):
                            borrowed_name = outbound_list[0][0].split('HTA')[0] + line[0]
                            inv_tuple = (str(quadifier(borrowed_name)), int(line[1]))
                            outbound_list.append(inv_tuple)
                        else:
                            inv_tuple = (str(line[0]), int(line[1]))
                            second_chance_list.append(inv_tuple)
                            continue
                    else:
                        inv_tuple = (str(quadifier(line[0])), int(line[1]))
                        outbound_list.append(inv_tuple)
        else:
            names_and_numbers = False

            # When checking film, this checks for film name and number elements.

            for line in reader:
                try:
                    if line[5]:
                        names_and_numbers = True
                except IndexError:
                    pass

                if after_string in str(line):
                    break

                if 'Helmet' in str(line):
                    type_string = 'Helmet'
                else:
                    type_string = 'Film'

                if 'HTA' in str(line[0]):
                    if line[0].split('HTA')[0] == '':
                        if len(outbound_list):
                            borrowed_name = outbound_list[0][0].split('HTA')[0] + line[0]
                            inv_tuple = (str(quadifier(borrowed_name)), int(line[1]), str(type_string))
                            outbound_list.append(inv_tuple)
                        else:
                            inv_tuple = (str(line[0]), int(line[1]), str(type_string))
                            second_chance_list.append(inv_tuple)
                            continue
                    else:
                        inv_tuple = (str(quadifier(line[0])), int(line[1]), str(type_string))
                        outbound_list.append(inv_tuple)

            # If there are any entries in the second chance list, this tries them again.

            if len(second_chance_list):
                for item in second_chance_list:
                    if len(outbound_list):
                        borrowed_name = outbound_list[0][0].split('HTA')[0] + item[0]
                        if count_film is False:
                            inv_tuple = (str(quadifier(borrowed_name)), int(item[1]))
                        else:
                            inv_tuple = (str(quadifier(borrowed_name)), int(item[1]), str(item[2]))
                        outbound_list.append(inv_tuple)

            # If there are names and numbers, this checks for an associated file in Sublimation Artwork 2.
            # If it finds one, it checks the size. If it's shorter than an art page box, it adds an entry to the
            # outbound list noting the presence of names and numbers as well as a path to the file. If it's taller,
            # placeholder entries may first be added to the list to ensure that it doesn't run off the page.

            if names_and_numbers:
                try:
                    sub_art_path = ''.join(('R:\\Sublimation Artwork 2\\', csv_path.split(os.path.sep)[-3], '\\', csv_path.
                                            split(os.path.sep)[-2]))
                    for file in os.listdir(sub_art_path):
                        if 'HTA' not in file:
                            names_and_numbers_path = ''.join((sub_art_path, '\\', str(file)))
                            illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
                            illustrator.Open(names_and_numbers_path)
                            names_and_numbers_document = illustrator.ActiveDocument
                            illustrator.ExecuteMenuCommand('selectall')
                            illustrator.ExecuteMenuCommand('group')
                            if names_and_numbers_document.Selection[0].Height <= 765:
                                outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                                names_and_numbers_document.Close(2)
                                break
                            else:
                                hta_box_count = len(outbound_list)
                                name_and_number_box_count = (names_and_numbers_document.Selection[0].Height // 765) + 1
                                if (hta_box_count + name_and_number_box_count) // 10 > hta_box_count // 10\
                                        and (hta_box_count + name_and_number_box_count) % 10 != 0:
                                    spacer_count = (((hta_box_count // 10) + 1) * 10) - hta_box_count
                                    while spacer_count:
                                        outbound_list.append((' ', 0, ' '))
                                        spacer_count -= 1
                                    names_and_numbers_document.Close(2)
                                    break
                                else:
                                    outbound_list.append(('Names and Numbers', 0, names_and_numbers_path))
                                    names_and_numbers_document.Close(2)
                                    break
                except FileNotFoundError:
                    pass

    return outbound_list


def quadifier(whole_hta_string):
    # This function makes sure that the HTA number has the proper number of digits, which is four.
    # It does this by adding or removing leading zeroes.
    # Now with 20% more not crashing on film pieces with no HTA!


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

    return whole_hta_string
