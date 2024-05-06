# Team_Shop_Automaton
A suite of tools for processing Wilson Team Shop orders. Probably only useful to me, but feel free to grab any chunks you find useful!

This project is now feature-complete! I'm still making little adjustments as needed.

## Main Components

### wts_order_processor.py
> Processes an entire WTS order, from inbox to outbox, creating all required documents and database entries along the way.

### wts_order_multiprocessor.py
> Runs the order processor until the inbox has been emptied.
> Also runs the inbox_highlanderer before it begins.

### inbox_highlanderer.py
> Removes all duplicate emails from the postorder inbox.

### folder_molder.py
> Creates folders for shop closure emails and places the .CSV order count inside of them.

### count_transmogrifier.py
> Parses the contents of a .CSV order count and returns a list of film, heat transfer, or raised helmet HTAs, along with their quantities and, in the case of film, their color palette type (film/helmet).
> Contains a filter which ensures proper, four-digit HTA numbers and catches some common typos.

### count_comparitron.py
> Takes a list of HTAs and quantities (from count_transmogrifier), compares them to the inventory database, and returns a list of not-in-stock HTAs and their ordering quantities.
> This one also has a filter for HTAs with a regrettable space in their name, which also checks for versions with the space replaced with a hyphen and with the spaces just removed entirely.

### art_page_combobulator.py
> Takes a count sheet and makes Order Art Pages and Film Art Pages from it.

### player_number_cruncher.py
> Takes a count sheet and makes Player Numbers Art Pages from it.

### ml_order_former.py
> Takes one or more Order Art Pages and makes an MLOrder.

### log_borg.py
> Parses the contents of an order folder and the .CSV order count contained therein, then inserts a record into the Heat Transfer Inventory database.

### wts_outbot.py
> Parses the contents of an order folder and the .CSV order count contained therein, then prepares one or more emails with appropriate text and attachments.

### decamessage_dispatcher.py
> Sends ten orders' worth of draft emails.

### draft_counter.py
> Returns the number of orders currently contained in the draft box.

## Manual Controllers

### combobulator_controller.py
> Takes a folder path and runs it through the art_page_combobulator.

### ml_order_former_controller.py
> Takes a folder path and runs it through the ml_order_former.

### number_cruncher_controller.py
> Takes a folder path and runs it through the player_number_cruncher.

### log_borg_controller.py
> Takes a folder path and runs it through the log_borg.

### outbot_controller.py
> Takes a folder path and runs it through the wts_outbot.

## Resources

### wilson_colors.py
> A library of Wilson CMYK colors for use with Illustrator.

### Improved Art Sheet.ai
> A streamlined art sheet template for use with the art sheet modules.

### Number Block Template.ai
> A holding space for the ml_order_former to create player number header blocks in, rather than opening the art pages to retrieve them.

### config.ini
> A configuration file, now required for almost all other modules. Keep it in the same file as all of the modules.

### gen_py_remove.bat
> A script that fixes the missing CLSID map bug. I'm still looking for a way to run this automatically that actually works when it is needed.

