# ---------------------------------------------------------------------------------------------------------------------#
# art_page_combobulator.py
#
# This module takes a list of HTAs and quantities (and sometimes film types) and makes an art sheet. This is
# accomplished through Illustrator's COM interface, the interaction with which through Python is almost completely
# undocumented, at least in terms that made any sense to me at the beginning of this. If you have any idea what you're
# doing with this kind of thing, or you don't and you'd like some help from someone very marginally familiar, please
# reach out to me!
#
# NOTE: This module works exclusively with the Improved Art Sheet, which is available on my GitHub, where you probably
# got this file. Attempting to use the regular art sheet will result in it not working at all, so do not do that.
#
# ---------------------------------------------------------------------------------------------------------------------#

import more_itertools
import os
import win32com.client

from configparser import ConfigParser
from pathlib import Path

import count_comparitron
import count_transmogrifier


def main(inbound_list, team_name, store_number, page_type, page_number):
    # inbound_list is a list of HTAs, ordering quantities, and sometimes film types.
    # team_name is the name of the shop.
    # store_number is the number of the shop.
    # page_type determines the type of page that is produced. Current options include "Film" and "Heat Transfer".
    # page_number is the page number for occasions where there are more than ten items. Zero indicates a single page.

    # This block opens the configuration file and retrieves the required folder/file paths.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)

    art_files_compiled_path = config['Folder Paths']['art_files_compiled_path']
    blank_sheet_path = config['Folder Paths']['blank_sheet_path']
    working_folder_path = config['Folder Paths']['working_folder_path']

    # This chunk opens Illustrator if it isn't open, then opens the blank art sheet and assigns it a name.

    illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
    illustrator.UserInteractionLevel = -1
    illustrator.Open(blank_sheet_path)
    art_sheet = illustrator.ActiveDocument

    # This block creates a header string and replaces the placeholder header on the art sheet with it.

    if page_type == 'Film':
        if page_number == 0:
            header = ''.join((team_name, '\n', store_number, '\nFilm Art Actual'))
        else:
            header = ''.join((team_name, '\n', store_number, '\nFilm Art Actual 0', str(page_number)))
    elif page_type == 'Heat Transfer':
        if page_number == 0:
            header = ''.join((team_name, '\n', store_number, '\nORDER Art Page'))
        else:
            header = ''.join((team_name, '\n', store_number, '\nORDER Art Page 0', str(page_number)))

    for item in art_sheet.TextFrames:
        if item.Contents == 'Fakeville Baseballers 19Q 0123456789':
            item.Contents = header
            break

    # box_count tracks into which box the next HTA will be placed.
    # tall_mode is activated when an HTA is taller than one of the boxes on the art sheet. When in Tall Mode,
    # vertical placement is based on the bottom of the preceding HTA for the remainder of the column.

    box_count = 0
    tall_mode = False

    # This loop opens each required HTA in turn and places it on the art sheet.

    for order_hta in inbound_list:
        # If the order has a lot of film names and numbers, they might need to be pushed to a new page.
        # This ends the current page if it hits one of the spacers used to do that.

        if order_hta[0] == ' ':
            break

        # Before doing all the HTA stuff, this checks if the item is a film Names and Numbers tag instead of an HTA.
        # If it is, it must be the last item in the list, and this block will run instead of going through the for loop
        # again. This will look for the film Names and Numbers file in the Sublimation Artwork 2 folder and, upon
        # finding it, will place its contents in one or more available blocks.

        if order_hta[0] == 'Names and Numbers':
            # This opens the file, the path to which has been stored in the list by the count_transmogrifier.
            # The contents of the file are selected, grouped, and copied, then the file is closed.

            illustrator.Open(order_hta[2])
            names_and_numbers_document = illustrator.ActiveDocument

            illustrator.ExecuteMenuCommand('selectall')
            illustrator.ExecuteMenuCommand('group')

            names_and_numbers_document.Copy()
            names_and_numbers_document.Close(2)

            art_sheet.Paste()

            # If the names and numbers are taller than one box and set to be placed at the bottom of the first column,
            # this bumps them to the top of the second.

            if box_count == 4 and art_sheet.Selection[0].Height > 765:
                box_count = 5

            # This finds the correct box and moves the names and numbers into it.

            for item in art_sheet.PathItems:
                if item.Name == 'R' + str(box_count):
                    rectangle = item
                    break

            art_sheet.Selection[0].Left = rectangle.Left + 20

            if tall_mode:
                art_sheet.Selection[0].Top = last_bottom
            else:
                art_sheet.Selection[0].Top = rectangle.Top - 20

            if art_sheet.Selection[0].Width <= rectangle.Width:
                art_sheet.Selection[0].Translate(DeltaX=(rectangle.Width * .25))

            # If they don't fit into a box, this creates a background rectangle for easy visibility.

            if art_sheet.Selection[0].Height > rectangle.Height or art_sheet.Selection[0].Width > rectangle.Width:
                background_top = art_sheet.Selection[0].Top + 10
                background_left = art_sheet.Selection[0].Left - 10
                background_width = art_sheet.Selection[0].Width + 20
                background_height = art_sheet.Selection[0].Height + 20

                background_rectangle = art_sheet.PathItems.Rectangle(background_top,
                                                                     background_left,
                                                                     background_width,
                                                                     background_height)

                background_rectangle.FillColor = rectangle.FillColor

                illustrator.ExecuteMenuCommand('sendToFront')
                background_rectangle.Selected = True
                illustrator.ExecuteMenuCommand('sendBackward')
            break

        # This bit opens the hta and assigns it a name.

        hta_path = ''.join((art_files_compiled_path, '\\', order_hta[0], '.ai'))
        illustrator.Open(hta_path)
        hta = illustrator.ActiveDocument

        # The next forty-ish lines select and remove any "Film", "Helmet", "UV Print", etc. labels from the top right
        # of the HTA, then replace it if needed.
        #
        # This took an astounding amount of time to get working correctly, thanks to inconsistencies in HTA production.

        if len(hta.TextFrames) >= 3:
            removal_list = []
            for item in hta.TextFrames:
                removal_list.append(item.Left)
            for item in hta.TextFrames:
                if item.Left == min(removal_list):
                    if item.Height < 100:
                        item.Delete()

        if page_type == 'Film':
            label_string = order_hta[2]
            left_list = []
            bottom_list = []

            for item in hta.PageItems:
                left_list.append(item.Left)
                bottom_list.append(item.Top - item.Height)

            far_left = min(left_list)
            top_bottom = max(bottom_list)

            label_height = 39.13232421875
            font_top = top_bottom + label_height

            film_label = hta.TextFrames.Add()
            film_label.Contents = label_string
            film_label.Top = font_top - (.5 * label_height)
            film_label.Left = far_left

            scarlet = win32com.client.Dispatch('Illustrator.CMYKColor')
            scarlet.Cyan = 0.000000
            scarlet.Magenta = 97.000003
            scarlet.Yellow = 74.000001
            scarlet.Black = 5.000000

            hta.CharacterStyles.RemoveAll()
            fresh_style = hta.CharacterStyles.Add('fresh_style')
            funky_attributes = fresh_style.CharacterAttributes

            funky_attributes.Size = 36
            funky_attributes.FillColor = scarlet

            fresh_style.ApplyTo(film_label.TextRange)

            hta.SelectObjectsOnActiveArtboard()
            film_label.TextRange.Select(True)
        elif page_type == 'Heat Transfer':
            hta.SelectObjectsOnActiveArtboard()

        # Next, the HTA is closed and pasted on the art sheet in an appropriate position.

        illustrator.ExecuteMenuCommand('group')
        hta.Copy()
        hta.Close(2)

        for item in art_sheet.PathItems:
            if item.Name == 'R' + str(box_count):
                rectangle = item
                break

        art_sheet.Paste()

        art_sheet.Selection[0].Position = rectangle.Position

        if art_sheet.Selection[0].Width < rectangle.Width:
            translate_x = .5 * (rectangle.Width - art_sheet.Selection[0].Width)
            art_sheet.Selection[0].Translate(translate_x, 0)

        illustrator.ExecuteMenuCommand('group')

        if tall_mode:
            art_sheet.Selection[0].Top = last_bottom

        # The quantity for the HTA is added.

        for item in art_sheet.TextFrames:
            if item.Name == 'Q' + str(box_count):
                item.Contents = str(order_hta[1])
                break

        # If the HTA is abnormally tall, Tall Mode is activated to avoid overlapping.

        if art_sheet.Selection[0].Height > rectangle.Height:
            tall_mode = True

        # If the end of the first column has been reached, Tall Mode is turned off.

        if box_count == 4:
            tall_mode = False

        # Now, if Tall Mode is active, the bottom of the HTA is located so that the next one can be placed correctly.

        if tall_mode:
            last_bottom = art_sheet.Selection[0].Top - art_sheet.Selection[0].Height

        # The box count is incremented.

        box_count += 1

    # Now, it's time to save. This took me all day to figure out, so I'm going to complain informatively.
    #
    # First, just saving as an .AI file is easy; SaveAs() just takes a filepath and works as expected. However, it saves
    # in the compatibility mode for the version of Illustrator you're using, while company policy dictates the use of
    # CS5 compatibility. Can you just add a compatibility mode as a parameter? No, you must create an ai_save_options
    # profile, which is the only kind of thing so far, apart from Illustrator itself, that must be created using
    # Dispatch(). I'm sure this makes sense to the informed but I had no idea.
    #
    # Another annoyance is that, according to the javascript guide (for sensible people using the right language),
    # compatibility modes are referred to by text strings. Not so here; they're enumerated, and undocumentedly so!
    # For the record, CS5 is 15, CS6 is 16, and CC is 17. After that, they started using years, which match up (for
    # example, Illustrator 2022 is version 22).
    #
    # Finally, to save a .PDF, you use the usual SaveAs() function, with pdf_save_options in place of the .AI ones.

    file_name = header.replace('\n', ' ')
    save_path = ''.join((working_folder_path, '\\', team_name, '\\', store_number, '\\', file_name, '.ai'))

    ai_save_options = win32com.client.Dispatch('Illustrator.IllustratorSaveOptions')
    ai_save_options.Compatibility = 15
    art_sheet.SaveAs(save_path, ai_save_options)

    save_path = save_path.replace('.ai', '.pdf')
    pdf_save_options = win32com.client.Dispatch('Illustrator.PDFSaveOptions')
    art_sheet.SaveAs(save_path, pdf_save_options)

    art_sheet.Close(2)


def combobulate(folder_path, skip_heat_transfers=False):
    # This function allows other modules to call the combobulator using just a folder path.
    # Optionally, it can also be instructed to skip the heat transfer page.

    # This block imports the MiLB and Direct Transfer lists from the config file.
    # milb_list is a list of Minor League Baseball teams.
    # For MiLB orders, items that would ordinarily be adorned with heat transfers get film instead.
    # direct_transfer_list is a list of teams that only use direct heat transfers.
    # For these, the heat transfer ordering process is skipped entirely.

    config_path = Path(__file__).parent.absolute().joinpath('config.ini')
    config = ConfigParser()
    config.read(config_path)

    milb_list = str(config['Special Teams']['milb_list']).split(',')
    direct_transfer_list = str(config['Special Teams']['direct_transfer_list']).split(',')

    # This block finds the .CSV count file in the folder, if there is one, then runs it through the
    # count_comparitron and count_transmogrifier to get accurate heat transfer and film lists.

    csv_path, film_list, transfer_list = '', '', ''

    for file in os.listdir(folder_path):
        if file.endswith('.csv') and 'Count' in file:
            csv_path = os.path.join(folder_path, file)
            if any(word in str(folder_path.split(os.path.sep)[-2]) for word in milb_list):
                transfer_list = []
                film_list = count_comparitron.compare_helmets(count_transmogrifier.count_milb_film(csv_path))
                break
            elif any(word in str(folder_path.split(os.path.sep)[-2]) for word in direct_transfer_list):
                transfer_list = []
                film_list = count_comparitron.compare_helmets(count_transmogrifier.count_film(csv_path))
            else:
                transfer_list = count_comparitron.compare_transfers(count_transmogrifier.count_heat_transfers(csv_path))
                film_list = count_comparitron.compare_helmets(count_transmogrifier.count_film(csv_path))
                break

    if csv_path == '':
        print('No count found!')
        return True

    if film_list == 'Names and numbers not found!':
        return 'Early'

    # This block gets the Team Name and Store Number.

    team_name = folder_path.split(os.path.sep)[-2]
    store_number = folder_path.split(os.path.sep)[-1]

    # This block attempts to create one or more heat transfer art sheets. If there are more than ten items,
    # it splits the list into chunks of ten and runs the art_page_combobulator multiple times.

    if not skip_heat_transfers:
        try:
            if transfer_list:
                if len(transfer_list) <= 10:
                    main(transfer_list, team_name, store_number, 'Heat Transfer', 0)
                    print('Heat transfer page created!')

                else:
                    transfer_iter = more_itertools.chunked(transfer_list, 10)
                    page_number = 1
                    for sublist in transfer_iter:
                        main(sublist, team_name, store_number, 'Heat Transfer', page_number)
                        page_number += 1
                    print('Heat transfer pages created!')
            else:
                print('No heat transfers required!')
        except:
            print('Heat transfer page(s) FAILED!')
            return True

    # This one does the same but for film.

    try:
        if film_list:
            if len(film_list) <= 10:
                main(film_list, team_name, store_number, 'Film', 0)
                print('Film page created!')

            else:
                film_iter = more_itertools.chunked(film_list, 10)
                page_number = 1
                for sublist in film_iter:
                    main(sublist, team_name, store_number, 'Film', page_number)
                    page_number += 1
                print('Film pages created!')
        else:
            print('No film required!')
    except:
        print('Film page(s) FAILED!')
        return True

    # If anything went wrong, the module should have already returned True. Otherwise, it will now return False.

    return False
