# ---------------------------------------------------------------------------------------------------------------------#
# wts_order_multiprocessor.py
#
# This module runs the wts_order_processor until it runs out of orders. It runs the inbox highlanderer first to weed
# out duplicate emails.
#
# I'll rename this when I think of a better name.
#
# ---------------------------------------------------------------------------------------------------------------------#

import inbox_highlanderer
import wts_order_processor

from time import sleep

inbox_highlanderer.main()

persist = True

while persist:
    persist = wts_order_processor.main()
    sleep(5)
