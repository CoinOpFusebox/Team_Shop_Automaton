# ---------------------------------------------------------------------------------------------------------------------#
# report_generator.py
#
# This module provides a text rundown of the requirements of each order in the working folder. Hopefully someday I can
# just feed this data directly to another automaton, but for now at least I don't have to write things down in a
# notebook like a caveman.
#
# ---------------------------------------------------------------------------------------------------------------------#

import csv
import os
import count_comparitron
import count_transmogrifier

# OPTIONS
#
# Set working_folder_path to the full path of whatever folder holds your order folders.

working_folder_path = r'C:\Users\fredricg\Downloads\Working Folders'

# This block checks each folder that exists somewhere within the working folder.
# Any folder with no other folders in it is processed.

for dirpath, dirnames, filenames in os.walk(working_folder_path):
    if not dirnames:
        print('\n\n\n' + dirpath.split(os.path.sep)[-2] + ' ' + dirpath.split(os.path.sep)[-1])
        csv_path = ''

        for file in os.listdir(dirpath):
            if file.endswith('.csv'):
                csv_path = os.path.join(dirpath, file)

                heat_transfer_list = count_comparitron.main(count_transmogrifier.main(csv_path, False))

                if heat_transfer_list:
                    print('\nHeat Transfers:')
                    print(heat_transfer_list)

                film_list = count_transmogrifier.main(csv_path, True)

                if film_list:
                    print('\nFilm:')
                    print(film_list)

                before_string = '\'HT Adult CID\''
                search_string = 'Player Number'
                player_numbers = False

                with open(csv_path, 'r', errors="ignore") as csv_count:
                    reader = csv.reader(csv_count)

                    for line in reader:
                        if before_string in str(line):
                            break

                    for line in reader:
                        if search_string in str(line):
                            player_numbers = True
                            break

                    if player_numbers:
                        print('\nPlayer Numbers')

                if not heat_transfer_list and not film_list and not player_numbers:
                    print('\nNo items needed!')

                break
