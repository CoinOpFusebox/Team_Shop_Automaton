# Team_Shop_Automaton
A suite of tools for processing Wilson Team Shop orders. Probably only useful to me, but feel free to grab any chunks you find useful!

This project is intended to streamline and optimize, to the furthest capacity allowed by my skill level and available time, my current job duties.
At present, it features the following modules:

team_shop_outbot.py
> Parses the contents of an order folder and the .CSV order count contained therein, then prepares one or more emails with appropriate text and attachments.

outbot_controller.py
> Provides a simple CLI which allows the user to manually call team_shop_outbot by providing the path for the desired order folder.
> Good for testing, batch_hopper is probably more practical for everyday use.

log_borg.py
> Parses the contents of an order folder and the .CSV order count contained therein, then inserts a record into the Heat Transfer Inventory database.

log_borg_controller.py
> Provides a simple CLI which allows the user to manually call log_borg by providing the path for the desired order folder.
> Again, batch_hopper is probably more useful.

batch_hopper.py
> Scans a folder (currently hardcoded) then runs each bottom-level folder through both team_shop_outbot and log_borg.
> Process a bunch of orders, then log and send all of them with one click. Happy days!
