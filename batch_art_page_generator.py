# ---------------------------------------------------------------------------------------------------------------------#
# batch_art_page_generator.py
#
# This module scans the working folder for all bottom-level folders within, then runs them through the
# art_page_combobulator.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import win32com.client

import art_page_combobulator

# OPTIONS
#
# Set working_folder_path to the full path of whatever folder holds your order folders.

working_folder_path = r'C:\Users\fredricg\Downloads\Working Folders'

# This list prints at the end, in case anything goes wrong and needs individual attention.

problem_list = []

# This block checks each folder that exists somewhere within the working folder.
# Any folder with no other folders in it is processed.

illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')

for dirpath, dirnames, filenames in os.walk(working_folder_path):
    problem = False
    if not dirnames:
        print(dirpath)
        problem = art_page_combobulator.combobulate(dirpath)

        if problem:
            problem_list.append(dirpath)

        for document in illustrator.Documents:
            document.Close(1)

# This prints a list of problems for the user's perusal.

if problem_list:
    print('The following folders had one or more failed art pages:\n')
    for item in problem_list:
        print(item)

illustrator.Quit()