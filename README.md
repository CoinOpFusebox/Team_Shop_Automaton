# Team_Shop_Automaton
A suite of tools for processing Wilson Team Shop orders. Probably only useful to me, but feel free to grab any chunks you find useful!

This project is intended to streamline and optimize, to the furthest capacity allowed by my skill level and available time, my current job duties.
At present, it features the following modules:

outbot_controller.py
> Provides a simple CLI which allows the user to manually call team_shop_outbot.py by providing the path for the desired order folder.
> May not be needed once more progress has been made.

team_shop_outbot.py
> Parses the contents of an order folder and the .CSV order count contained therein, then prepares one or more emails with appropriate text and attachments.
