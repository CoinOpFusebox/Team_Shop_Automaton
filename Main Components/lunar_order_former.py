# ---------------------------------------------------------------------------------------------------------------------#
# lunar_order_former.py
#
# This module creates an MLOrder from an ORDER Art Page. It does not yet handle player numbers.
#
# ---------------------------------------------------------------------------------------------------------------------#

import os
import win32com.client

# OPTIONS
#
# Set blank_order_path to the path where you keep your blank MLOrder.

blank_order_path = r'C:\Users\fredricg\OneDrive - Amer Sports\Documents\Work Stuff\Templates\MLOrder.xlsm'


def main(folder_path):

    team_name = folder_path.split(os.path.sep)[-2]
    store_number = folder_path.split(os.path.sep)[-1]

    # art_sheet_path is the path for the current ORDER Art Page. This will change repeatedly in multi-page applications.
    # order_path is the path where the completed MLOrder will be saved.
    # multi_mode is activated when multiple art pages are present.
    # persevere will get switched off when the HTAs run out, causing the loop to end.
    # multi_page_count tracks the current number of MLOrder pages.

    art_sheet_path = ''.join([folder_path, '\\', team_name, ' ', store_number, ' ORDER Art Page.ai'])
    order_path = ''.join([folder_path, '\\', 'MLOrder ', team_name, ' ', store_number, '.xlsm'])

    multi_mode = False
    persevere = True
    multi_page_count = 1

    # This block opens Illustrator and Excel, as well as a blank MLOrder.

    illustrator = win32com.client.gencache.EnsureDispatch('Illustrator.Application')
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    order_sheet = excel.Workbooks.Open(blank_order_path)
    order_sheet_page = order_sheet.Worksheets('Order Form')

    # This block will first try to open a stand-alone ORDER Art Page. If it fails, it will look for one ending in "01".
    # If it finds a post-numbered art page, it will switch on multi_mode and continue.
    # Otherwise, the module ceases its function.

    try:
        illustrator.Open(art_sheet_path)
        art_sheet = illustrator.ActiveDocument
    except:
        try:
            art_sheet_path = ''.join([folder_path, '\\', team_name, ' ', store_number, ' ORDER Art Page 0',
                                      str(multi_page_count), '.ai'])
            illustrator.Open(art_sheet_path)
            art_sheet = illustrator.ActiveDocument
            multi_mode = True
        except:
            print('No art sheet found!')
            return

    # header_text holds the text that will be inserted at the top of the first page of an MLOrder.

    header_text = ''.join([team_name.rsplit(' ', 2)[-3], '\n', team_name.split()[-2], ' ', team_name.split()[-1], '\n',
                            store_number])

    while persevere:
        # This block ungroups everything on the art page so that the text blocks are selectable without their art.
        # There is probably a better way to do this.

        art_sheet.SelectObjectsOnActiveArtboard()
        illustrator.ExecuteMenuCommand('ungroup')
        illustrator.ExecuteMenuCommand('ungroup')
        illustrator.ExecuteMenuCommand('deselectall')

        # box_count tracks the number of HTA boxes used on the current MLOrder page.

        box_count = 0

        # text_frame_list holds a list of Text Frames in the art page.
        # This is necessary in order to reverse the list as a TextFrames object is not reversible.
        # Reversing the list is necessary because of the way the art_page_combobulator inserts HTAs.

        text_frame_list = []
        for item in art_sheet.TextFrames:
            text_frame_list.append(item)

        # This block checks each TextFrame on the art page.
        #
        # If the TextFrame is named "Header" it is used as a vessel for the header text, copied, pasted
        # onto the MLOrder, and resized to fit the box. This only happens on the first page of each MLOrder.
        #
        # If the text contains the string "HTA", box_count is incremented and the text is copied and pasted into the
        # next available box on the MLOrder, then resized to fit the box.

        for item in reversed(text_frame_list):
            if multi_page_count % 3 == 1:
                if item.Name == 'Header':
                    item.Contents = header_text
                    item.Copy()
                    order_sheet_page.Range('B7').Select()
                    header_height = excel.Selection.Height
                    header_width = excel.Selection.Width
                    order_sheet_page.PasteSpecial(Format='Bitmap')
                    excel.Selection.Height = header_height
                    if excel.Selection.Width > header_width:
                        excel.Selection.Width = header_width
                    continue
            if 'HTA' in item.Contents:
                box_count += 1
                illustrator.ExecuteMenuCommand('deselectall')
                item.TextRange.Select(True)
                art_sheet.Selection[0].Copy()
                item.TextRange.DeSelect()
                if box_count < 6:
                    order_sheet_page.Range('D' + str(21 + 3 * box_count)).Select()
                else:
                    order_sheet_page.Range('M' + str(21 + 3 * (box_count - 5))).Select()
                box_height = excel.Selection.Height
                box_width = excel.Selection.Width
                order_sheet_page.PasteSpecial(Format='Bitmap')
                excel.Selection.Height = box_height
                if excel.Selection.Width > box_width:
                    excel.Selection.Width = box_width
                continue

        # quantity_index tracks which quantity box is to be copied next.

        quantity_index = 0

        # Using the box_count to track how many quantities are present, this loop sets each quantity on the MLOrder to
        # match the ones on the art page.

        while quantity_index < box_count:
            for item in art_sheet.TextFrames:
                if item.Name == 'Q' + str(quantity_index):
                    if quantity_index < 5:
                        order_sheet_page.Range('B' + str(24 + 3 * quantity_index)).Select()
                        excel.Selection.Value = item.Contents
                        quantity_index += 1
                        break
                    else:
                        order_sheet_page.Range('K' + str(24 + 3 * (quantity_index - 5))).Select()
                        excel.Selection.Value = item.Contents
                        quantity_index += 1
                        break

        # This closes the art sheet.

        art_sheet.Close(1)

        # If multi_mode is active, the multi_page_count is incremented to signify the passage to the next page.
        # The art_sheet_path is updated so that the next page can be opened. If another art page is not found, persevere
        # is deactivated and the loop ends.

        if multi_mode:
            multi_page_count += 1

            art_sheet_path = ''.join([folder_path, '\\', team_name, ' ', store_number, ' ORDER Art Page 0',
                                      str(multi_page_count), '.ai'])

            try:
                illustrator.Open(art_sheet_path)
                art_sheet = illustrator.ActiveDocument

            except:
                persevere = False

            # If the next art page has been opened, this loop gets the next MLOrder page ready.
            #
            # Since each MLOrder only has three pages, a new MLOrder is required every three pages. The order_path is
            # updated to add "Part ##" to then end, the MLOrder is saved and closed, and a new one is opened. The
            # order_path is then updated again to ensure that, if this loop is not activated again, the final MLOrder is
            # named correctly.
            #
            # If there are pages remaining on the current MLOrder, the next one is activated.

            if persevere:
                if multi_page_count % 3 == 1:
                    order_book_count = multi_page_count // 3
                    order_path = ''.join([folder_path, '\\', 'MLOrder ', team_name, ' ', store_number, ' Part 0',
                                          str(order_book_count) , '.xlsm'])
                    order_sheet.SaveAs(order_path)
                    excel.DisplayAlerts = False
                    excel.Quit()
                    order_path = ''.join([folder_path, '\\', 'MLOrder ', team_name, ' ', store_number, ' Part 0',
                                          str(order_book_count + 1), '.xlsm'])
                    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
                    order_sheet = excel.Workbooks.Open(blank_order_path)
                    order_sheet_page = order_sheet.Worksheets('Order Form')

                else:
                    order_sheet_page = order_sheet.Worksheets('Page ' +
                                                              str(multi_page_count - (3 * ((multi_page_count - 1) // 3))))
                    order_sheet_page.Activate()


        else:
            persevere = False

    # Once the art pages have run out, the MLOrder is saved and Excel is closed.

    order_sheet.SaveAs(order_path)
    excel.DisplayAlerts = False
    excel.Quit()

    print('MLOrder generated!')
