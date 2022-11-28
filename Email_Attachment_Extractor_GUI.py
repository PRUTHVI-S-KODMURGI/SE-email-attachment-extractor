import imaplib # module to connect to the imap server
import email # module to retrieve emails
import os # module to save the attachments
from tkinter import * # module for the GUI interface
from datetime import datetime # module to retrieve the current timestamp
from tkinter import messagebox

# Initializing a Tkinter window and giving it a title followed by giving the window a size with the help of the geometry method
root = Tk()
root.title("email-attachment-extractor")
root.geometry("1920x1080")
root.configure(bg='grey')

# Heading of the Tkinter window
Heading = LabelFrame(root, bd=2, relief="groove", bg="light yellow")
Heading.place(x=0, y=0, width=1920, height=55)
app_name = Label(Heading, text="Email Attachment Extractor", font="arial 20 bold italic", bg="light pink", fg="blue").grid(row=0, column=0, padx=600, pady=5)

global username, password, folder_path, delete, verbose, verbose_txt_area, no_emails

# Input for the Username
username_label = Label(root, text="Username: ",font='arial 12')
username_label.place(x=50, y=160)
username = StringVar()
username_entry = Entry(root, textvariable=username, width=50, font='arial 12')
username_entry.place(x=220, y=160)

# Input for the Password
password_label = Label(root, text="Password: ",font='arial 12')
password_label.place(x=50, y=210)
password = StringVar()
password_entry = Entry(root, textvariable=password, show='*', width=50, font='arial 12')
password_entry.place(x=220, y=210)

# Input for the output folder path
folder_path_label = Label(root, text="Output Folder Path: ",font='arial 12')
folder_path_label.place(x=50, y=260)
folder_path = StringVar()
folder_path_entry = Entry(root, textvariable=folder_path, width=50, font='arial 12')
folder_path_entry.place(x=220, y=260)

# Input for the Verbose
verbose_label = Label(root, text="Verbose: ",font='arial 12')
verbose_label.place(x=50, y=310)
verbose = StringVar(None, '0')
verbose_radio_yes = Radiobutton(root, text='Yes',font='arial 12', value='1', variable=verbose)
verbose_radio_yes.place(x=540, y=310)
verbose_radio_no = Radiobutton(root, text='No',font='arial 12', value='0', variable=verbose)
verbose_radio_no.place(x=595, y=310)

# Radio Button for selecting if emails need to get deleted after extracting the attachments
delete_label = Label(root, text="Delete the emails after extracting the attachments: ", font='arial 12')
delete_label.place(x=50, y=360)
delete = StringVar(None, '0')
delete_radio_yes = Radiobutton(root, text='Yes',font='arial 12', value='1', variable=delete)
delete_radio_yes.place(x=540, y=360)
delete_radio_no = Radiobutton(root, text='No',font='arial 12', value='0', variable=delete)
delete_radio_no.place(x=595, y=360)

# Input for the number of emails to be extracted
no_emails_label = Label(root, text="Number of Emails from which attachments need to get extracted: ",font='arial 12')
no_emails_label.place(x=50, y=410)
no_emails = IntVar()
no_emails_entry = Entry(root, textvariable=no_emails, width=10, font='arial 12')
no_emails_entry.place(x=540, y=410)

# Verbose displaying widget
verbose_area = LabelFrame(root, bd=2, relief="groove")
verbose_area.place(x=850, y=150, width=600, height=400)
bill_title = Label(verbose_area, text="VERBOSE", font="arial 15 bold", bd=4, relief="groove").pack(fill=X)
scroll_y = Scrollbar(verbose_area, orient=VERTICAL)
verbose_txt_area = Text(verbose_area, yscrollcommand=scroll_y.set)
scroll_y.pack(side=RIGHT, fill=Y)
scroll_y.config(command=verbose_txt_area.yview)
verbose_txt_area.pack(fill=BOTH, expand=1)

def clear_all():
    # Setting the various Entry and Text widgets to their default values
    username.set('')
    password.set('')
    folder_path.set('')
    verbose.set('0')
    delete.set('0')
    no_emails.set(0)
    verbose_txt_area.delete("1.0", "end")

def fetchattachments(username_val, password_val, folder_val, no_emails_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    verbose_txt_area.delete("1.0", "end")

    # User Authentication
    try:
        connection.login(username_val, password_val)
    except:
        messagebox.showerror("Invalid Credentials", "Incorrect Username or Password")
        clear_all()
        return
    
    try:
        (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

        (email_status, [emails_number]) = connection.search(None, 'ALL')
        emails = emails_number.split()
        no_emails_traverse = min(len(emails), no_emails_val)

        for em in range(len(emails)-1, len(emails)-1-no_emails_traverse, -1):
            dummy_email = emails[em]
            connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
            (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
            email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
            mail = email.message_from_string(email_body.decode('utf8'))

            if mail.is_multipart():
                compteur = 0
                for part in mail.walk():
                    compteur += 1

                    boolfileName = part.get_filename()
                    if bool(boolfileName): # None => False
                        # Creating a file name by appending the timestamp so that the file names become unique
                        timestamp = datetime.now()
                        append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                        fileName = part.get_filename().split('.')[0] + "_" + append_filename + "." + part.get_filename().split('.')[-1]
                        filePath = os.path.join(folder_val, fileName)

                        if not os.path.isfile(filePath): # Check if there is already a file with the same name (very rare)
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

        connection.close()
        connection.logout()
    except:
        messagebox.showerror("Error", "An error has occurred, please try again!")
        clear_all()
        return
    
    messagebox.showinfo("Extraction Completion Confirmation", "The required number of attachments have been successfully extracted from the specified number of emails.")

def fetchattachments_verbose(username_val, password_val, folder_val, no_emails_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    verbose_txt_area.delete("1.0", "end")

    try:
        connection.login(username_val, password_val)
    except:
        messagebox.showerror("Invalid Credentials", "Incorrect Username or Password")
        clear_all()
        return
    
    try:
        (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

        verbose_txt_area.insert(END, "There are "+str(int(emails_number[0]))+" emails inside "+username_val+" account from the imap server 'imap.gmail.com'")

        (email_status, [emails_number]) = connection.search(None, 'ALL')
        emails = emails_number.split()
        no_emails_traverse = min(len(emails), no_emails_val)

        for em in range(len(emails)-1, len(emails)-1-no_emails_traverse, -1):
            dummy_email = emails[em]
            verbose_txt_area.insert(END, '\nWorking on mail number '+dummy_email.decode('utf8')) # dummy_email is in binary mode

            connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
            (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
            email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
            mail = email.message_from_string(email_body.decode('utf8'))

            if mail.is_multipart():
                compteur = 0
                for part in mail.walk():
                    compteur += 1
                    verbose_txt_area.insert(END, '\nPart '+str(compteur)+' in mail.walk with Content Type : '+str(part.get_content_type))

                    boolfileName = part.get_filename()
                    if bool(boolfileName): # None => False
                        timestamp = datetime.now()
                        append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                        fileName = part.get_filename().split('.')[0] + "_" + append_filename + "." + part.get_filename().split('.')[-1]
                        verbose_txt_area.insert(END, '\nDetected file : '+fileName)
                        filePath = os.path.join(folder_val, fileName)

                        if not os.path.isfile(filePath): # Check if there is already a file with the same name
                            verbose_txt_area.insert(END, '\nWriting file')
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

        connection.close()
        connection.logout()
    except Exception:
        messagebox.showerror("Error", "An error has occurred, please try again!")
        clear_all()
        return
    
    messagebox.showinfo("Extraction Completion Confirmation", "The required number of attachments have been successfully extracted from the specified number of emails.")

def fetchattachments_delete(username_val, password_val, folder_val, no_emails_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    verbose_txt_area.delete("1.0", "end")

    try:
        connection.login(username_val, password_val)
    except:
        messagebox.showerror("Invalid Credentials", "Incorrect Username or Password")
        clear_all()
        return

    try:
        (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

        (email_status, [emails_number]) = connection.search(None, 'ALL')
        emails = emails_number.split()
        no_emails_traverse = min(len(emails), no_emails_val)

        for em in range(len(emails)-1, len(emails)-1-no_emails_traverse, -1):
            dummy_email = emails[em]
            connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
            (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
            email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
            mail = email.message_from_string(email_body.decode('utf8'))

            if mail.is_multipart():
                compteur = 0
                for part in mail.walk():
                    compteur += 1
                    boolfileName = part.get_filename()

                    if bool(boolfileName): # None => False
                        timestamp = datetime.now()
                        append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                        fileName = part.get_filename().split('.')[0] + "_" + append_filename + "." + part.get_filename().split('.')[-1]
                        filePath = os.path.join(folder_val, fileName)

                        if not os.path.isfile(filePath): # Check if there is already a file with the same name
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

        connection.close()
        connection.logout()
    except:
        messagebox.showerror("Error", "An error has occurred, please try again!")
        clear_all()
        return
    
    messagebox.showinfo("Extraction Completion Confirmation", "The required number of attachments have been successfully extracted from the specified number of emails.")

def fetchattachments_verbose_delete(username_val, password_val, folder_val, no_emails_val):
    """
    Connect to the mailox
    Check emails
    Save attachments
    """
    connection = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    verbose_txt_area.delete("1.0", "end")

    try:
        connection.login(username_val, password_val)
    except:
        messagebox.showerror("Invalid Credentials", "Incorrect Username or Password")
        clear_all()
        return

    try:
        (email_status, emails_number) = connection.select() # connect to the default folder (INBOX)

        verbose_txt_area.insert(END, "There are "+str(int(emails_number[0]))+" emails inside "+username_val+" account from the imap server 'imap.free.fr'")

        (email_status, [emails_number]) = connection.search(None, 'ALL')
        emails = emails_number.split()
        no_emails_traverse = min(len(emails), no_emails_val)

        for em in range(len(emails)-1, len(emails)-1-no_emails_traverse, -1):
            dummy_email = emails[em]
            verbose_txt_area.insert(END, '\nWorking on mail number '+dummy_email.decode('utf8')) # dummy_email is in binary mode

            connection.store(dummy_email, '+FLAGS', '(%s)' % 'SEEN')
            (type_email, msg_data) = connection.fetch(dummy_email, '(RFC822)')
            email_body = msg_data[0][1] # Retrieve all mail (same as mail source)
            mail = email.message_from_string(email_body.decode('utf8'))

            if mail.is_multipart():
                compteur = 0
                for part in mail.walk():
                    compteur += 1
                    verbose_txt_area.insert(END, '\nPart '+str(compteur)+' in mail.walk with Content Type : '+str(part.get_content_type))
                    
                    boolfileName = part.get_filename()
                    if bool(boolfileName): # None => False
                        timestamp = datetime.now()
                        append_filename = str(timestamp.day)+str(timestamp.month)+str(timestamp.year)+str(timestamp.hour)+str(timestamp.minute)+str(timestamp.second)
                        fileName = part.get_filename().split('.')[0] + "_" + append_filename + "." + part.get_filename().split('.')[-1]
                        verbose_txt_area.insert(END, '\nDetected file : '+fileName)
                        filePath = os.path.join(folder_val, fileName)

                        if not os.path.isfile(filePath): # Check if there is already a file with the same name
                            verbose_txt_area.insert(END, '\nWriting file')
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

            connection.store(dummy_email, '+FLAGS', '\\Deleted')

        connection.expunge()
        verbose_txt_area.insert(END, '\nEmails deleted')

        connection.close()
        connection.logout()
    except:
        messagebox.showerror("Error", "An error has occurred, please try again!")
        clear_all()
        return
    
    messagebox.showinfo("Extraction Completion Confirmation", "The required number of attachments have been successfully extracted from the specified number of emails.")

def select_fetch():
    verbose_val = int(verbose.get())
    delete_val = int(delete.get())

    try:
        dummy_var = int(no_emails.get())
    except:
        messagebox.showerror("Invalid Input", "The number of emails field should have an integer value")
        return
    
    if dummy_var <= 0:
        messagebox.showinfo("Invalid Input", "Please provide a positive integer to the number of emails field")
        return
    
    if verbose_val==0 and delete_val==0:
        fetchattachments(username.get(), password.get(), folder_path.get(), no_emails.get())
    elif verbose_val==1 and delete_val==0:
        fetchattachments_verbose(username.get(), password.get(), folder_path.get(), no_emails.get())
    elif verbose_val==0 and delete_val==1:
        fetchattachments_delete(username.get(), password.get(), folder_path.get(), no_emails.get())
    else:
        fetchattachments_verbose_delete(username.get(), password.get(), folder_path.get(), no_emails.get())

# Button to start the extraction function
extract_button = Button(root, text="Extract",font='arial 15 bold',fg='green', command=select_fetch)
extract_button.place(x=320, y=500)

if __name__ == "__main__":
    root.mainloop()