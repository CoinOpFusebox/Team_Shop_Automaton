# ---------------------------------------------------------------------------------------------------------------------#
# number_cruncher_controller.py
#
# This module provides a crude means of interfacing with the player_number_cruncher module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import player_number_cruncher

while True:
    folder_path = input("Enter a filepath to generate player numbers. Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()

    else:
        if os.path.isdir(folder_path):
            player_number_cruncher.main(folder_path)
        else:
            print('Invalid path!\n')
