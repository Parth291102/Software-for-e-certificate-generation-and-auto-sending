import smtplib
import os

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#Setup SMTP port and server name

smtp_port = 587                 #Standard secure SMTP port 
smtp_server = "smtp.gmail.com"  #Google SMTP Server

email_from = "sl2ecde55@gmail.com"
email_list = ["20bec081@nirmauni.ac.in"]
pswd = "xasaxcnmgxchsrqb"
subject = "Result Announcement!" 

folder_path = "./SL_assignment/"

#To access each file in folder_path use a for loop
def send_mails(email_list):
    for filename in os.listdir(folder_path):
        if filename.endswith(".jpg"):
            file_path = os.path.join(folder_path, filename)
        for person in email_list:
            
            #Made the body of email 
            body = f"""
            Greetings of the Day!!!

            The Organization appreciates all the participants 
            for actively participating in the event. We extend
            our heartfelt Congratulations to all the winners. 
            We look forward to meeting you again at our events.            
            
            Winners can find their certificates in the attachment.  
            """
            #make a MIME object to define parts of the email
            msg = MIMEMultipart()
            msg['From'] = email_from
            msg['To'] = person
            msg['Subject'] = subject

            #Attach the body of the message 
            msg.attach(MIMEText(body, 'plain'))
                                       
            #Open the file in python as binary form
            if os.path.isfile(file_path):
                with open(file_path, "rb") as attachment:    #r means read and b means binary
                    part = MIMEApplication(attachment.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)   # Lastly we are attaching to the msg object

            #Convert message to string
            text = msg.as_string()  # This line will fire the email 

            #Connect with the server 
            print("Connecting to server...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email_from, pswd)
            print("Successfully connected to the Server :)")
            
            #Send emails to person
            print(f"Sending email to: {person} ...")
            server.sendmail(email_from, person, text)
            print(f"Email sent to: {person}")
            server.quit()
            
send_mails(email_list)