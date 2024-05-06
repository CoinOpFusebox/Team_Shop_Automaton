# ---------------------------------------------------------------------------------------------------------------------#
# outbot_controller.py
#
# This module provides a crude means of interfacing with the wts_outbot module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import wts_outbot

while True:
    folder_path = input("Enter a filepath to generate email(s). Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()

    else:
        if os.path.isdir(folder_path):
            wts_outbot.main(folder_path)
        else:
            print('Invalid path!\n')
