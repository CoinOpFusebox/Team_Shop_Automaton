# ---------------------------------------------------------------------------------------------------------------------#
# log_borg_controller.py
#
# This module provides a crude means of interfacing with the log_borg module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import log_borg

while True:
    folder_path = input("Enter a filepath to generate a database entry. Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()

    else:
        if os.path.isdir(folder_path):
            log_borg.main(folder_path)
        else:
            print('Invalid path!\n')
