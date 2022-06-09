# Team_Shop_Automaton
A suite of tools for processing Wilson Team Shop orders. Probably only useful to me, but feel free to grab any chunks you find useful!

This project is intended to streamline and optimize, to the furthest capacity allowed by my skill level and available time, my current job duties.
At present, it features the following modules:

team_shop_outbot.py
> Parses the contents of an order folder and the .CSV order count contained therein, then prepares one or more emails with appropriate text and attachments.

outbot_controller.py
> Provides a simple CLI which allows the user to manually call team_shop_outbot by providing the path for the desired order folder.
> Good for testing, batch_caber_tosser is probably more practical for everyday use.

log_borg.py
> Parses the contents of an order folder and the .CSV order count contained therein, then inserts a record into the Heat Transfer Inventory database.

log_borg_controller.py
> Provides a simple CLI which allows the user to manually call log_borg by providing the path for the desired order folder.
> Again, batch_caber_tosser is probably more useful.

batch_caber_tosser.py
> Scans a folder (currently hardcoded) then runs each bottom-level folder through both team_shop_outbot and log_borg.
> Process a bunch of orders, then log and send all of them with one click. Happy days!

count_transmogrifier.py
> Parses the contents of a .CSV order count and returns a list of either film or heat transfer HTAs, along with their quantities and, in the case of film, their color palette (film/helmet).
> Contains a filter which ensures proper, four-digit HTA numbers.

count_comparitron.py
> Takes a list of HTAs and quantities (say, from count_transmogrifier), compares them to the inventory database, and returns a list of not-in-stock HTAs and their ordering quantities.
> This one also has a filter for HTAs with a regrettable space in their name, which also checks for versions with the space replaced with a hypen and with the spaces just removed entirely.

report_generator.py
> Uses count_transmogrifier and count_comparitron to generate a text report featuring the heat transfers, film, and player numbers (for now, just the presence or absence of numbers, presented without further information) required by each order in the working folder.
> Crude, but far quicker than the notebooks I've been using for the past the past three years.
