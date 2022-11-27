import imaplib # module to connect to the imap server
import email # module to retrieve emails
import os # module to save the attachments
# import argparse # module to parse the arguments
# from time import sleep # module to give a pause to the execution
from tkinter import * # module for the GUI interface
import datetime # module to retrieve the current timestamp

# Initializing a Tkinter window and giving it a title followed by giving the window a size with the help of the geometry method
root = Tk()
root.title("email-attachment-extractor")
root.geometry("1920x1080")

# Heading of the Tkinter window
Heading = LabelFrame(root, bd=2, relief="groove", bg="light yellow")
Heading.place(x=0, y=0, width=1920, height=55)
app_name = Label(Heading, text="Email Attachment Extractor", font="arial 20 bold italic", bg="light pink", fg="blue").grid(row=0, column=0, padx=600, pady=5)

global username, password, folder_path, delete, verbose

# Input for the Username
username_label = Label(root, text="Username: ")
username_label.place(x=30, y=100)
username = StringVar()
username_entry = Entry(root, textvariable=username, width=50)
username_entry.place(x=150, y=100)

# Input for the Password
password_label = Label(root, text="Password: ")
password_label.place(x=30, y=150)
password = StringVar()
password_entry = Entry(root, textvariable=password, show='*', width=50)
password_entry.place(x=150, y=150)

# Input for the output folder path
folder_path_label = Label(root, text="Output Folder Path: ")
folder_path_label.place(x=30, y=200)
folder_path = StringVar()
folder_path_entry = Entry(root, textvariable=folder_path, width=50)
folder_path_entry.place(x=150, y=200)

# Input for the Verbose
verbose_label = Label(root, text="Verbose: ")
verbose_label.place(x=30, y=250)
verbose = StringVar(None, '0')
verbose_radio_yes = Radiobutton(root, text='Yes', value='1', variable=verbose)
verbose_radio_yes.place(x=300, y=250)
verbose_radio_no = Radiobutton(root, text='No', value='0', variable=verbose)
verbose_radio_no.place(x=350, y=250)

# Radio Button for selecting if emails need to get deleted after extracting the attachments
delete_label = Label(root, text="Delete the emails after extracting the attachments: ")
delete_label.place(x=30, y=300)
delete = StringVar(None, '0')
delete_radio_yes = Radiobutton(root, text='Yes', value='1', variable=delete)
delete_radio_yes.place(x=300, y=300)
delete_radio_no = Radiobutton(root, text='No', value='0', variable=delete)
delete_radio_no.place(x=350, y=300)

def clear_all():
    username.set('')
    password.set('')
    folder_path.set('')
    verbose.set('0')
    delete.set('0')

def fetchattachments(username_val, password_val, folder_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.free.fr', 993)
    connection.login(username_val, password_val)
    (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

    (email_status, [emails_number]) = connection.search(None, 'ALL')
    for dummy_email in emails_number.split():

        connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
        (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
        email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
        mail = email.message_from_string(email_body.decode('utf8'))

        if mail.is_multipart():
            compteur = 0
            for part in mail.walk():
                compteur += 1
                
                timestamp = datetime.now()
                append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                fileName = part.get_filename() + "_" + append_filename
                if bool(fileName): # None => False
                    filePath = os.path.join(folder_val, fileName)
                    if not os.path.isfile(filePath): # Check is there is already a file with the same name
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()

    connection.close()
    connection.logout()

def fetchattachments_verbose(username_val, password_val, folder_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.free.fr', 993)
    connection.login(username_val, password_val)
    (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

    print("\nThere is {} emails inside {} account from the imap server {}".format(int(emails_number[0]), username_val, 'imap.free.fr'))

    (email_status, [emails_number]) = connection.search(None, 'ALL')
    for dummy_email in emails_number.split():
        print('\n *** Working on mail number {}'.format(dummy_email.decode('utf8'))) # dummy_email is in binary mode

        connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
        (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
        email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
        mail = email.message_from_string(email_body.decode('utf8'))

        if mail.is_multipart():
            compteur = 0
            for part in mail.walk():
                compteur += 1
                print('Part {} in mail.walk with Content Type : {}'.format(compteur, part.get_content_type()))
                
                timestamp = datetime.now()
                append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                fileName = part.get_filename() + "_" + append_filename
                if bool(fileName): # None => False
                    print('Detected file : {}'.format(fileName))
                    filePath = os.path.join(folder_val, fileName)
                    if not os.path.isfile(filePath): # Check is there is already a file with the same name
                        print('Writing file')
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()

    connection.close()
    connection.logout()

def fetchattachments_delete(username_val, password_val, folder_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.free.fr', 993)
    connection.login(username_val, password_val)
    (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

    (email_status, [emails_number]) = connection.search(None, 'ALL')
    for dummy_email in emails_number.split():

        connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
        (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
        email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
        mail = email.message_from_string(email_body.decode('utf8'))

        if mail.is_multipart():
            compteur = 0
            for part in mail.walk():
                compteur += 1
                
                timestamp = datetime.now()
                append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                fileName = part.get_filename() + "_" + append_filename
                if bool(fileName): # None => False
                    filePath = os.path.join(folder_val, fileName)
                    if not os.path.isfile(filePath): # Check is there is already a file with the same name
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()

    connection.close()
    connection.logout()

def fetchattachments_verbose_delete(username_val, password_val, folder_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.free.fr', 993)
    connection.login(username_val, password_val)
    (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

    print("\nThere is {} emails inside {} account from the imap server {}".format(int(emails_number[0]), username_val, 'imap.free.fr'))

    (email_status, [emails_number]) = connection.search(None, 'ALL')
    for dummy_email in emails_number.split():
        print('\n *** Working on mail number {}'.format(dummy_email.decode('utf8'))) # dummy_email is in binary mode

        connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
        (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
        email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
        mail = email.message_from_string(email_body.decode('utf8'))

        if mail.is_multipart():
            compteur = 0
            for part in mail.walk():
                compteur += 1
                print('Part {} in mail.walk with Content Type : {}'.format(compteur, part.get_content_type()))
                
                timestamp = datetime.now()
                append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                fileName = part.get_filename() + "_" + append_filename
                if bool(fileName): # None => False
                    print('Detected file : {}'.format(fileName))
                    filePath = os.path.join(folder_val, fileName)
                    if not os.path.isfile(filePath): # Check is there is already a file with the same name
                        print('Writing file')
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
        connection.store(dummy_email, '+FLAGS', '\\Deleted')

    connection.expunge()
    print('\nEmails deleted')

    connection.close()
    connection.logout()

def select_fetch():
    verbose_val = int(verbose.get())
    delete_val = int(delete.get())
    if verbose_val==0 and delete_val==0:
        fetchattachments(username.get(), password.get(), folder_path.get())
    elif verbose_val==1 and delete_val==0:
        fetchattachments_verbose(username.get(), password.get(), folder_path.get())
    elif verbose_val==0 and delete_val==1:
        fetchattachments_delete(username.get(), password.get(), folder_path.get())
    else:
        fetchattachments_verbose_delete(username.get(), password.get(), folder_path.get())

# Button to start the extraction function
extract_button = Button(root, text="Extract", command=select_fetch)
extract_button.place(x=60, y=350)

if __name__ == "__main__":
    root.mainloop()