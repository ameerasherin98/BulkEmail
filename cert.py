import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders



i = 0

email_user = 'ieee.siae@gmail.com'
email_password = 'siaeieeeks'
email_send =  open(r'C:\Users\DELL\Desktop\Blog\.vscode\certificates\emails.txt').readlines()
filename = open(r'C:\Users\DELL\Desktop\Blog\.vscode\certificates\pdf.txt').readlines()                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 


while i<100:
        subject = 'IEEE KS - APJAKTU Webinar | Certificate'
        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject
        msg.preamble = 'Multipart massage.\n'

        html="""<!doctype html>
<html>
  <body>
  <p>Greetings from IEEE Kerala Section</p>
  <p>Thanking you for being part of webinar series organised by <b>Strategic Initiatives for Academic Engagement</b>, IEEE Kerala Section in association with APJ Abdul Kalam Technological University, Kerala </p>
  <p>You can find your certificate along with this e-mail</p>
  <p>Regards</p>

<p style="color:grey">Dr Sameer SM<br>
Chair - Strategic Initiative for Academic Engagement | IEEE KS.<br>
Professor - Dept. of ECE | National Institute of Technology, Calicut.<br>
Head - Centre for Training & Placement | National Institute of Technology, Calicut.<br></p>
  </body>
</html>"""

        part2=MIMEText(html,'html')
        msg.attach(part2)
        #body = 
       

        #msg.attach(MIMEText(body,'plain'))
        filename1 = filename[i].rstrip()
        print(filename1)
        attachment = open(filename1,'rb')
        print(attachment)
        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',"attachment; filename= "+filename1)

        msg.attach(part)
        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(email_user,email_password)

        server.sendmail(email_user,email_send[i],text)
        print ("Email Sent!!")
        del filename1
        del attachment
        print (i)
        server.quit()
        i = i+1