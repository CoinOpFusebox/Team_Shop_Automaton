# ---------------------------------------------------------------------------------------------------------------------#
# wts_order_multiprocessor.py
#
# This module runs the wts_order_processor until it runs out of orders. It also runs the inbox highlanderer and preorder
# data logger before it starts running orders.
#
# I'll rename this when I think of a better name.
#
# ---------------------------------------------------------------------------------------------------------------------#

import inbox_highlanderer
import preorder_data_logger
import wts_order_processor

from time import sleep

inbox_highlanderer.main()
preorder_data_logger.main()

persist = True

while persist:
    persist = wts_order_processor.main()
    sleep(5)
