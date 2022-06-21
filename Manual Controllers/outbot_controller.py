# ---------------------------------------------------------------------------------------------------------------------#
# outbot_controller.py
#
# This module provides a crude means of interfacing with the team_shop_outbot module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import team_shop_outbot

while True:
    folder_path = input("Enter a filepath to generate email(s). Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()

    else:
        if os.path.isdir(folder_path):
            team_shop_outbot.main(folder_path)
        else:
            print('Invalid path!\n')
