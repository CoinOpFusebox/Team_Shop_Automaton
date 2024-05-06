# ---------------------------------------------------------------------------------------------------------------------#
# ml_order_former_controller.py
#
# This module provides a crude means of interfacing with the ml_order_former module.
# It may be rendered obsolete later as greater automation is implemented.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os

import ml_order_former
import player_number_cruncher

while True:
    folder_path = input("Enter a filepath to generate any required MLOrders. Enter \"exit\" to exit.\n")

    if folder_path == 'exit':
        exit()
    else:
        if os.path.isdir(folder_path):
            hta_count = ml_order_former.order_heat_transfers(folder_path)
            ml_order_former.order_player_numbers(folder_path, player_number_cruncher.main(folder_path), hta_count)
        else:
            print('Invalid path!\n')
