# ---------------------------------------------------------------------------------------------------------------------#
# combobulator_controller.py
#
# This module provides a crude means of interfacing with the art_page_combobulator module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os

import art_page_combobulator

while True:
    folder_path = input("Enter a filepath to generate any required art pages. Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()
    else:
        if os.path.isdir(folder_path):
            art_page_combobulator.combobulate(folder_path)
        else:
            print('Invalid path!\n')
