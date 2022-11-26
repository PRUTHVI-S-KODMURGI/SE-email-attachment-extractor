import imaplib # module to connect to the imap server
import email # module to retrieve emails
import os # module to save the attachments
import argparse # module to parse the arguments
from time import sleep # module to give a pause to the execution
from tkinter import * # module for the GUI interface

# Initializing a Tkinter window and giving it a title followed by giving the window a size with the help of the geometry method
root = Tk()
root.title("email-attachment-extractor")
root.geometry("1920x1080")

# Heading of the Tkinter window
Heading = LabelFrame(root, bd=2, relief="groove", bg="light yellow")
Heading.place(x=0, y=0, width=1920, height=55)
app_name = Label(Heading, text="Email Attachment Extractor", font="arial 20 bold italic", bg="light pink", fg="blue").grid(row=0, column=0, padx=600, pady=5)

# Input for the Username
username_label = Label(root, text="Username: ")
username_label.place(x=30, y=100)
username = StringVar()
username_entry = Entry(root, textvariable=username)
username_entry.place(x=150, y=100)

# Input for the Password
password_label = Label(root, text="Password: ")
password_label.place(x=30, y=150)
password = StringVar()
password_entry = Entry(root, textvariable=password, show='*')
password_entry.place(x=150, y=150)

# Input for the output folder path
folder_path_label = Label(root, text="Output Folder Path: ")
folder_path_label.place(x=30, y=200)
folder_path = StringVar()
folder_path_entry = Entry(root, textvariable=folder_path)
folder_path_entry.place(x=150, y=200)

# Input for the Imap Server
# imap_label = Label(root, text="IMAP Server: ")
# imap_label.place(x=30, y=250)
# imap_server = StringVar()
# imap_server_entry = Entry(root, textvariable=imap_server)
# imap_server_entry.place(x=150, y=200)

# Radio Button for selecting if emails need to get deleted after extracting the attachments
delete_label = Label(root, text="Delete the emails after extracting the attachments: ")
delete_label.place(x=30, y=300)
delete = 0
delete_radio_yes = Radiobutton(root, text='Yes', value=1, variable=delete)
delete_radio_yes.place(x=300, y=300)
delete_radio_no = Radiobutton(root, text='No', value=0, variable=delete)
delete_radio_no.place(x=350, y=300)


def fetchattachments(username, password, folder, verbose = True, delete = False, imapserver = 'imap.free.fr'):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL(imapserver, 993)
    connection.login(username, password)
    (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)
    if verbose:
        print("\nThere is {} emails inside {} account from the imap server {}".format(int(emails_number[0]), username, imapserver))

    (email_status, [emails_number]) = connection.search(None, 'ALL')
    for dummy_email in emails_number.split():
        if verbose:
            print('\n *** Working on mail number {}'.format(dummy_email.decode('utf8'))) # dummy_email is in binary mode
        connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
        (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
        email_body = msg_data[0][1] #retrieve all mail (same as mail source)
        mail = email.message_from_string(email_body.decode('utf8'))

        if mail.is_multipart():
            compteur = 0
            for part in mail.walk():
                compteur += 1
                if verbose:
                    print('Part {} in mail.walk with Content Type : {}'.format(compteur, part.get_content_type()))
                fileName = part.get_filename() #Renvoie None si pas de fichier attaché
                if bool(fileName): # None => False
                    if verbose:
                        print('Detected file : {}'.format(fileName))
                    filePath = os.path.join(folder, fileName)
                    if not os.path.isfile(filePath): # Check is there is already a file with the same name
                        if verbose:
                            print('Writing file')
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
        if delete:
            connection.store(dummy_email, '+FLAGS', '\\Deleted')
    if delete:
        connection.expunge()
        if verbose:
            print('\nEmails deleted')
    connection.close()
    connection.logout()

#button to start the extraction function
b = Button(root,text="Extract",command= fetchattachments)
b.pack()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download all attachments from an email account")
    parser.add_argument('folder', help="the folder where you want to save your attachements")
    parser.add_argument('email', type=str, help="imap account user id")
    parser.add_argument('password', type=str, help="imap account password")
    parser.add_argument('-d', '--delete', action='store_true', help="delete emails after saving attachment")
    parser.add_argument('-v', '--verbose', action='store_true', help="verbose mode")
    parser.add_argument('-imap', help="imap server (by default is imap.free.fr)")
    args = parser.parse_args()
    folder = args.folder
    identifiant = args.email
    password = args.password
    if args.verbose:
        verbose = True
        print("\nVerbose mode activated\n")
    else:
        verbose = False
    if args.delete:
        delete = True
        print(f"\nMessages will be deleted from {identifiant}\nYou have now 5 seconds to cancel")
        sleep(5)
        print("\nOK, let's go")
    else:
        delete = False
    if args.imap:
        imap = args.imap
        print(identifiant)
        print(password)
        fetchattachments(str(identifiant), str(password), folder, verbose, delete, imap)
    else:
        fetchattachments(identifiant, password, folder, verbose, delete)
    root.mainloop()