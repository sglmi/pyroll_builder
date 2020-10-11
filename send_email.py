# import smtplib

# from string import Template

# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText


# EMAIL_HOST = "smtp.gmail.com"
# EMAIL_PORT = 587
# EMAIL_HOST_USER = "saeidtempmail@gmail.com"
# EMAIL_HOST_PASSWORD = "123passhello"


# def get_contacts(filename):
#     """
#     Return two lists names, emails containing names and email addresses
#     read from a file specified by filename.
#     """

#     names = []
#     emails = []
#     with open(filename, mode="r", encoding="utf-8") as contacts_file:
#         for a_contact in contacts_file:
#             names.append(a_contact.split()[0])
#             emails.append(a_contact.split()[1])
#     return names, emails


# def read_template(filename):
#     """
#     Returns a Template object comprising the contents of the
#     file specified by filename.
#     """

#     with open(filename, "r", encoding="utf-8") as template_file:
#         template_file_content = template_file.read()
#     return Template(template_file_content)


# def main():
#     names, emails = get_contacts("mycontacts.txt")  # read contacts
#     message_template = read_template("message.txt")

#     # set up the SMTP server
#     s = smtplib.SMTP(host=EMAIL_HOST, port=EMAIL_PORT)
#     s.starttls()
#     s.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)

#     # For each contact, send the email:
#     for name, email in zip(names, emails):
#         msg = MIMEMultipart()  # create a message

#         # add in the actual person name to the message template
#         message = message_template.substitute(PERSON_NAME=name.title())

#         # Prints out the message body for our sake
#         print(message)

#         # setup the parameters of the message
#         msg["From"] = EMAIL_HOST_USER
#         msg["To"] = email
#         msg["Subject"] = "This is TEST"

#         # add in the message body
#         msg.attach(MIMEText(message, "plain"))

#         # send the message via the server set up earlier.
#         s.send_message(msg)
#         del msg

#     # Terminate the SMTP session and close the connection
#     s.quit()


# if __name__ == "__main__":
#     main()

from Tkinter import *


def donothing():
    filewin = Toplevel(root)
    button = Button(filewin, text="Do nothing button")
    button.pack()


root = Tk()
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=donothing)
filemenu.add_command(label="Open", command=donothing)
filemenu.add_command(label="Save", command=donothing)
filemenu.add_command(label="Save as...", command=donothing)
filemenu.add_command(label="Close", command=donothing)

filemenu.add_separator()

filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Undo", command=donothing)

editmenu.add_separator()

editmenu.add_command(label="Cut", command=donothing)
editmenu.add_command(label="Copy", command=donothing)
editmenu.add_command(label="Paste", command=donothing)
editmenu.add_command(label="Delete", command=donothing)
editmenu.add_command(label="Select All", command=donothing)

menubar.add_cascade(label="Edit", menu=editmenu)
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=donothing)
helpmenu.add_command(label="About...", command=donothing)
menubar.add_cascade(label="Help", menu=helpmenu)

root.config(menu=menubar)
root.mainloop()