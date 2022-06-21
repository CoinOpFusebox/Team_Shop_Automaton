# ---------------------------------------------------------------------------------------------------------------------#
# count_transmogrifier.py
#
# This module takes a .CSV count file and returns a list of either heat transfers or film and their quantities.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import re


def main(csv_path, count_film):
    # csv_path contains the path for the count sheet.
    # If count_film is True, film HTAs will be counted. An additional entry will be added to each tuple to denote type.

    # outbound_list is a list that will contain a tuple for each HTA, count, and sometimes film type.
    # If it remains empty, the order has none of the selected type of item.

    outbound_list = []

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

        if count_film is False:
            for line in reader:
                if after_string in str(line):
                    break

                inv_tuple = (str(quadifier(line[0])), int(line[1]))
                outbound_list.append(inv_tuple)
        else:
            for line in reader:
                if after_string in str(line):
                    break

                if 'Helmet' in str(line):
                    type_string = 'Helmet'
                else:
                    type_string = 'Film'

                inv_tuple = (str(quadifier(line[0])), int(line[1]), str(type_string))
                outbound_list.append(inv_tuple)
    return outbound_list


def quadifier(whole_hta_string):
    # This function makes sure that the HTA number has the proper number of digits, which is four.
    # It does this by adding or removing leading zeroes.
    # Now with 20% more not crashing on film pieces with no HTA!

    try:
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
    except:
        return 'N/A'
    
