# ---------------------------------------------------------------------------------------------------------------------#
# inbox_highlanderer.py
#
# This module removes all duplicate emails from the postorder inbox.
#
# ---------------------------------------------------------------------------------------------------------------------#
import win32com.client


def main():
    # This block opens Outlook and defines the required folders.

    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    inbox = mapi.GetDefaultFolder(6).Folders['Postorders']
    archive = mapi.Folders.Item(3).Folders['Archive']

    # This gets a list of all emails in the Postorders folder and sorts them in chronological order, oldest-first.

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", Descending=False)

    # This terminates the module if there are no emails.

    if not len(messages):
        # print('No emails!')
        return

    # This block checks the subject list for the current email's subject, archives the email if it's a duplicate, and
    # adds its subject to the list if it isn't.

    subject_list = []

    for item in messages:
        if item.Subject in subject_list:
            item.Move(archive)
        else:
            subject_list.append(item.Subject)


main()
