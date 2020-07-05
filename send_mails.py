import smtplib
import pandas as pd
import csv
import ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart  
from email.mime.base import MIMEBase  
from email import encoders 

sender_email = "abhisinghdeveloper7@gmail.com"
sender_name = "Abhishek Singh"
password = "your password"

e = pd.read_csv("test_mail.csv")
receiver_emails = e['receiver_emails'].values
receiver_names = e["receiver_names"].values

for receiver_email, receiver_name in zip(receiver_emails, receiver_names):
        
        
        print("Sending to " + receiver_name)
        msg = MIMEMultipart()
        msg['Subject'] = 'Advance Testing | ' + receiver_name + ' | Dont Worry !! Dont Panic'
        msg['From'] = formataddr((sender_name, sender_email))
        msg['To'] = formataddr((receiver_name, receiver_email))
        
        msg.attach(MIMEText("""<h2>Hello again, """ + receiver_name + """</h2>

                            <p> Hello, 

                            Greetings from Build with LetsUpgrade!!</p>

                          <br>

                            <p>We hope you are doing well during the COVID-19 Pandemic. While the whole nation is on lockdown, we are excited to cheer you up a little.</p>

                          <br>

                            <p>Congratulations on your successful registration on the Website!! We are glad you joined our community. </p>

                         <br>

                        <p>Here’s a quick introduction – Build with LetsUpgrade is a 2 month-long program conducted by LetsUpgrade with an aim to help beginners get started with Open Source Development. </p>

                        <br>

                        <br>

                        <p>As you already know that we have kick-started our Build with LetsUpgrade Project  Registrations from the 27th of June. Further, we wish to keep you informed with the developments regarding the BWLU Program and we’ll update them on all our social media handles throughout the program.</p>

                           <br>

                        <p>Building of Projects begins from 25th July 2020. This is the best opportunity to grow your skills, towards Open-Source contribution.</p>

                            <br>

                        <p>As such, make sure to follow all our social media accounts and get ready to upgrade your knowledge and skills.

                        </p>

                            <br>

                            <p>For Detailed Schedule:</p>

                          

                            <a href=  ' https://docs.google.com/document/d/1a5NBcuH1Bpubpx9Lki2QpuE2BoX2kEzU7Dw7wmx0JVY/edit?usp=drivesdk  '>Detailed Schedule for Build With LetsUpgrade</a>

                               <br>

                            <p> We would like you all, to fill up the given form provided in the link to get along this amazing journey.</p>

                                 <a href= 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/de_JXblXPe'> Registration form</a>

                                <p>Kindly fill out the form with all the correct details.</p>

                                 <br>

                                   <p>Those who filled the form, need not fill it again!</p>

                                 <br>


                                  <p>Selections for the Participants, Mentors, and Managers will be done on the basis of the responses given in the above form.</p>

                                   <br>

                                     <p>After the successful selection of the projects Mentors and Managers,  Only selected participants will be invited with another Google Form to choose the projects in which they want to contribute.</p>

                                <br>

                                  <p>We are eagerly looking forward to seeing you contribute.</p>        

                                  <br>

                                   <p>Till then, stay safe and keep learning and keep following our social media accounts.</p>

                                     <br>

                                     <p>STAY TUNED!</p>

                                    <br>

                                      <p>Join the LetsUpgrade Community Group for more: </p>

                                    <a href = 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/4zWJ4WyCO_'>LetsUpgrade community</a>          

                                     <br>

                                      <p>📱 Follow us on our Social Media handles to get more updates and winning swags:</p>

                                    <ul>                          

                                  <li><a href= 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/tb5W4_MlHk'>Website</a></li>
                                <li><a href = 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/tb5W4_MlHk' >Instagram</a></li>
                               <li><a href= 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/JiHTUDZO3w '>Twitter</a></li>
                                <li><a href = 'https://mail.letsupgrade.in/links/Cr3J8NEBJ/vguvh2z31/yz42F-OQxF/bws1MIMoBZ' >Facebook</a></li>
                                         


                                   <br>

                                     <p>Regards,</p>

                                      <p>Build with LetsUpgrade</p>              

                                  """, 'html'))


        filename = "Detailed_Schedule_new.rar"

        try:
            with open(filename, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
            )

            msg.attach(part)
        except Exception as e:
                print(f'Oh no! We didnt found the attachment!n{e}')
                break

        try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, msg.as_string())
                print('Email sent!')
        except Exception as e:
                print(f'Oh no! Something bad happened!n{e}')
                break
        finally:
                print('Closing the server...')
                server.quit()