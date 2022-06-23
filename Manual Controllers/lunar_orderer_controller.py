# ---------------------------------------------------------------------------------------------------------------------#
# lunar_orderer_controller.py
#
# This module provides a crude means of interfacing with the lunar_order_former module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os

import lunar_order_former

while True:
    folder_path = input("Enter a filepath to generate any required MLOrders. Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()
    else:
        if os.path.isdir(folder_path):
            lunar_order_former.main(folder_path)
        else:
            print('Invalid path!\n')
