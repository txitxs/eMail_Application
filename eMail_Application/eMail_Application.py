"""
+================================================+
|                                                |
|                                                |
|                                                |
|     __           .__   __                      |
|   _/  |_ ___  ___|__|_/  |_ ___  ___  ______   |
|   \   __\\  \/  /|  |\   __\\  \/  / /  ___/   |
|    |  |   >    < |  | |  |   >    <  \___ \    |
|    |__|  /__/\_ \|__| |__|  /__/\_ \/____  >   |
|                \/                 \/     \/    |
|                                                |
|                                                |
|                                                |
+================================================+
"""
#Dependencies for the project
from tkinter import *
import smtplib
import re
from pwinput import pwinput
from email.utils import parseaddr
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os 

#erstellt das GUI Window
window = Tk() 
window.geometry("1024x768")
window.title("Email Automation Application")
icon = PhotoImage(file="mail-inbox-app.png")
window.iconphoto(True, icon)

#regex for email validation
regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"  

# Funktion zum Senden der E-Mail
def send_email(senderEmail, senderPassword, receiverEmails, subject, messageBody, 
               attachmentPaths, emailProvider, emailProviderPortNumber):
    #Connect to SMTP server
 
    #Dynamically choose email provider
    smtpServer = f"smtp.{emailProvider}"
    
    #Dynamically choose port number
    smtp_port = emailProviderPortNumber
    
    server = smtplib.SMTP(smtpServer, smtp_port)
    server.starttls()
    server.login(senderEmail, senderPassword)
    
    for receiverEmail in receiverEmails:
        #Creates the message u want to send out
        
        msg = MIMEMultipart()
        msg["From"] = senderEmail
        msg["To"] = receiverEmail
        msg["Subject"] = subject
        
        #Adds message bodys
        msg.attach(MIMEText(messageBody, "plain"))
        
        #Add all attachments to the Email you want to sends
        for attachmentPath in attachmentPaths:
            if os.path.isfile(attachmentPath):
                with open(attachmentPath, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(attachmentPath)}')
                msg.attach(part)
        else:
                result_label.config(text=f"Attachment file not found: {attachmentPath}")
                return
            
        #Send email
        server.sendmail(senderEmail, receiverEmail,msg.as_string())
        
    #Close connection to the running server
    server.quit()

def check(email):
    email = email.strip()  
    if re.fullmatch(regex, email):
        print(f"The email address \"{email}\" you entered is valid.")
        return True  
    else:
        print(f"The email address \"{email}\" is invalid.")
        return False 

def submit_email_info():
    senderEmail = senderEmailEntry.get()
    if not check(senderEmail):  
        result_label.config(text="Invalid sender email address.")  
        return 


#User Input for all the Variables you need to enter for sending the email
#senderEmail = check(input("Please enter your email address: "))
emailPassword = pwinput("Please enter your password: ").strip
emailProvider = input("Please enter your email provider: ").strip
emailBody = input("Please enter the message body: ")

while True:
    try:
        emailProviderPortNumber = input("Please enter the port number for the email provider: ")
        emailProviderPortNumberInt = int(emailProviderPortNumber)
        break
    except ValueError:
         print("Invalid input. Please enter a valid Number for the port.")



def checkReceiverEmails(receiverEmails):
    while True:
        if re.fullmatch(regex, receiverEmails):
            print(f"The email address \"{receiverEmails}\" you entered is valid.")
            return receiverEmails
        else:
            receiverEmails = input("Invalid Email. Please enter a valid email address: ")

#Additional inputs for multiple recipients and attachments
receiverEmails = checkReceiverEmails(input("Please enter the email addresses of the recipients (separated by commas:  ,  ): "))
receiverEmails = [email.strip() for email in receiverEmails.split(",")]
subject = input("Please enter the subject of the email: ")
attachmentPaths = input("Please enter the file paths of the attachments (separated by commas): ")
attachmentPaths = [path.strip() for path in attachmentPaths.split(",")]

#Sends the email out
send_email(senderEmail, emailPassword, receiverEmails, subject,
           emailBody, attachmentPaths, emailProvider, emailProviderPortNumberInt)

senderEmailLaber = Label(window, text = "Sender Email:")
senderEmailLaber.pack()
senderEmailEntry = Entry(window, width=50)
senderEmailEntry.pack()

passwordLabel = Label(window, text= "Password:")
passwordLabel.pack()
passwordEntry = Entry(window, width=50, show='*')
passwordEntry.pack()

receiverEmailsLabel = Label(window, text="Receiver Emails (separated by commas):")
receiverEmailsLabel.pack()
receiverEmailsEntry = Entry(window, width=50)
receiverEmailsEntry.pack()

subjectLabel = Label(window, text="Subject:")
subjectLabel.pack()
subjectEntry = Entry(window, width=50)
subjectEntry.pack()

bodyLabel = Label(window, text="Email Body:")
bodyLabel.pack()
bodyEntry = Text(window, height=10, width=50)
bodyEntry.pack()

attachmentPathsLabel = Label(window, text="Attachment Paths (separated by commas):")
attachmentPathsLabel.pack()
attachmentPathsEntry = Entry(window, width=50)
attachmentPathsEntry.pack()

emailProviderLabel = Label(window, text="Email Provider (e.g., gmail.com):")
emailProviderLabel.pack()
emailProviderEntry = Entry(window, width=50)
emailProviderEntry.pack()

portLabel = Label(window, text="Port Number:")
portLabel.pack()
portEntry = Entry(window, width=50)
portEntry.pack()

# Label für die Ergebnisanzeige
result_label = Label(window, text="")
result_label.pack()

# Button zum Absenden der Email
sendButton = Button(window, text="Send Email", command=submit_email_info)
sendButton.pack()

# Main loop starten
window.mainloop()