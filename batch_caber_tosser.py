# ---------------------------------------------------------------------------------------------------------------------#
# batch_caber_tosser.py
#
# This module scans the working folder for all bottom-level folders within, then runs them through the team_shop_outbot
# and log_borg modules.
#
# This sort of makes outbot_controller and log_borg_controller obsolete for daily use.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import log_borg
import team_shop_outbot

# OPTIONS
#
# Set working_folder_path to the full path of whatever folder holds your order folders.

working_folder_path = r'C:\Users\fredricg\Downloads\Working Folders'

# This list prints at the end, in case anything goes wrong and needs individual attention.
# This is superior to the old plan, which was crashing.

problem_list = []

# This lonely block checks each folder that exists somewhere within the working folder.
# Any folder with no other folders in it is processed.

for dirpath, dirnames, filenames in os.walk(working_folder_path):
    if not dirnames:
        print(dirpath)
        try:
            team_shop_outbot.main(dirpath)
        except:
            problem_list.append('One or more emails have failed to send.\n')
        try:
            log_borg.main(dirpath)
        except:
            problem_list.append('One or more records have not been properly inserted.\n')

if problem_list:
    print(problem_list)
