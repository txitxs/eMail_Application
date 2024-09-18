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
 
    try:
        
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
    
        #Email was sent succesfule
        result_label.config(text="Email seccessfully sent!")
        
    except Exception as error: 
            result_label.config(text=f"Error: {error}")

def check(email):
    email = email.strip()  
    return re.fullmatch(regex, email)

def submit_email_info():
    senderEmail = senderEmailEntry.get()
    emailPassword = passwordEntry.get()
    receiverEmails = receiverEmailsEntry.get().split(",")
    subject = subjectEntry.get()
    emailBody = bodyEntry.get("1.0", END)
    attachmentPaths = attachmentPathsEntry.get().split(",")
    emailProvider = emailProviderEntry.get().strip()

    #validate port number
    try:
        emailProviderPortNumberInt = int(portEntry.get())
    except ValueError:
        result_label.config(text="Invalid port number. Please enter a valid number.")
        return

#Sends the email out
    send_email(senderEmail, emailPassword, receiverEmails, subject,
                emailBody, attachmentPaths, emailProvider, emailProviderPortNumberInt)

senderEmailLabel = Label(window, text = "Sender Email:")
senderEmailLabel.pack()
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