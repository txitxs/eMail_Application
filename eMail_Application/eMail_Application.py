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

#erstellt das GUI Window
window = Tk() 
window.geometry("1024x768")
window.title("Email Automation Application")
window.iconphoto(file="")
window.mainloop() 

#Variables
regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"

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
        msg["From"] = send_email
        msg["To"] = receiverEmail
        msg["subject"] = subject
        
        #Adds message body
        msg.attach(MIMEText(messageBody, "plain"))
        
        #Add all attachments to the Email you want to sends
        for attachmentPath in attachmentPaths:
            with open(attachmentPath, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {attachmentPath}')
            msg.attach(part)
            
        #Send email
        server.sendmail(senderEmail, receiverEmail,msg.as_string())
        
    #Close connection to the running server
    server.quit()

def check(senderEmail):
    while True:
        senderEmail = senderEmail.strip()
        if re.fullmatch(regex, senderEmail):
            print(f"The email address \"{senderEmail}\" you entered is valid.")
            return senderEmail
        else:
            senderEmail = input("Invalid Email. Please enter a valid email address: ")


#User Input for all the Variables you need to enter for sending the email
senderEmail = check(input("Please enter your email address: ").strip)
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