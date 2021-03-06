# Team_Shop_Automaton
A suite of tools for processing Wilson Team Shop orders. Probably only useful to me, but feel free to grab any chunks you find useful!

This project is intended to streamline and optimize, to the furthest capacity allowed by my skill level and available time, my current job duties.
At present, it features the following modules:

## Main Components

### preorder_data_logger.py *New!*
> Pasrses the preorder emails for accurate and (hopefully) consistent folder naming.

### folder_molder.py *New!*
> Creates folders for shop closure emails and places the .CSV order count inside of them.

### count_transmogrifier.py
> Parses the contents of a .CSV order count and returns a list of either film or heat transfer HTAs, along with their quantities and, in the case of film, their color palette (film/helmet).
> Contains a filter which ensures proper, four-digit HTA numbers.

### count_comparitron.py
> Takes a list of HTAs and quantities (say, from count_transmogrifier), compares them to the inventory database, and returns a list of not-in-stock HTAs and their ordering quantities.
> This one also has a filter for HTAs with a regrettable space in their name, which also checks for versions with the space replaced with a hypen and with the spaces just removed entirely.

### art_page_combobulator.py
> Takes a count sheet and makes ORDER Art Pages and Film Art Actual pages from it.

### lunar_order_former.py
> Takes one or more ORDER Art Pages and makes an MLOrder.

### log_borg.py
> Parses the contents of an order folder and the .CSV order count contained therein, then inserts a record into the Heat Transfer Inventory database.

### team_shop_outbot.py
> Parses the contents of an order folder and the .CSV order count contained therein, then prepares one or more emails with appropriate text and attachments.

## Batch Processing

### report_generator.py
> Uses count_transmogrifier and count_comparitron to generate a text report featuring the heat transfers, film, and player numbers (for now, just the presence or absence of numbers, presented without further information) required by each order in the working folder.
> Crude, but far quicker than the notebooks I've been using for the past the past three years.

### batch_art_page_generator.py
> Checks each base-level file for a count sheet and runs them all through the art_page_combobulator.

### batch_lunar_orderer.py
> Checks each base-level file for a count sheet and runs them all through the lunar_order_former.

### batch_caber_tosser.py
> Scans a folder (currently hardcoded) then runs each bottom-level folder through both team_shop_outbot and log_borg.
> Process a bunch of orders, then log and send all of them with one click. Happy days!

## Manual Controllers

### combobulator_controller.py
> Takes a folder path and runs it through the art_page_combobulator.

### lunar_orderer_controller.py
> Takes a folder path and runs it through the lunar_order_former.

### log_borg_controller.py
> Provides a simple CLI which allows the user to manually call log_borg by providing the path for the desired order folder.
> Again, batch_caber_tosser is probably more useful.

### outbot_controller.py
> Provides a simple CLI which allows the user to manually call team_shop_outbot by providing the path for the desired order folder.
> Good for testing, batch_caber_tosser is probably more practical for everyday use.

## Resources

### wilson_colors.py
> A library of Wilson CMYK colors for use with Illustrator.

### Improved Art Sheet.ai
> A streamlined art sheet template for use with the art sheet modules.

### TeamShopDB.db *New!*
> A simple database that holds paired team names and shop IDs. Required for preoder_data_logger and folder_molder.

### config.ini *New!*
> A configuration file, now required for almost all other modules. Keep it in the same file as all of the modules.
