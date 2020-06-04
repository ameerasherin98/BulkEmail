import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import os

pre=os.path.dirname(os.path.realpath(__file__))
fname='test.xlsx'
path=os.path.join(pre,fname)
e= pd.read_excel (path)
emails = e['Email'].values


sender_email = "ieee.siae2@gmail.com"
for email in emails:
    receiver_email = email
    password = "siaeieeeks"

    message = MIMEMultipart("alternative")
    message["Subject"] = "IEEE Kerala Section - APJAKTU Webinar 3 Login Credentials"
    message["From"] = sender_email
    message["To"] = receiver_email

    # Create the plain-text and HTML version of your message
    text = """\
            Greetings from IEEE Kerala Section!!.

Thank you for showing interest towards the Webinar series on Accreditation & Quality of Engineering Education on ABET perspective.

The Strategic Initiatives for Academic Engagement, IEEE Kerala Section in association with APJ Abdul Kalam Technological University (APJAKTU), Kerala is back with the third and  final webinar of the webinar series on ABET Accreditation Criteria by an eminent and esteemed resource person - Dr SK Ramesh Director - AIMS2 Program; Professor of Electrical and Computer Engineering - California State University, Northridge, USA

Registration is free and is open to all academicians. Fresh registration is mandatory to attend the webinar even if you have registered for the previous webinars. 
Logon to https://bit.ly/AccrABETWebinar and join us on 22 May, 2020 from 08:00 PM - 09:30 PM IST on Zoom. Find the flyer here : https://bit.ly/W3Flyer

Thank you for being a part of our previous Webinars on Accreditation and Quality Improvement in Engineering Education- An ABET Perspective. We are constantly striving to provide the ideal experience for our delegates, and your inputs help us to define that experience. 
Thus we request you to fill the below survey latest by 21st May, 2020. https://bit.ly/AccrWSFeedback

Looking forward to your presence during the Webinar.

Regards

Dr Sameer SM
Chair - Strategic Initiative for Academic Engagement | IEEE KS.
Professor - Dept. of ECE | National Institute of Technology, Calicut
Head - Centre for Training & Placement | National Institute of Technology, Calicut


    """
    html = """\
    <html>
    <body>
        <p><i>Greetings from IEEE Kerala Section!</i></p>
<br>
<p>You are receiving this mail as you have registered to attend the <b>Webinar on “Continuous Improvement and ABET Accreditation Process” </b>of the Webinar Series on - Accreditation and Quality Improvement in Engineering Education, organised by the<i> Strategic Initiative for Academic Engagements</i>, IEEE Kerala Section in association with APJ Abdul Kalam Technological University, Kerala (APJAKTU).</p>

<p style="color:red">Please find the participation link for the meeting at the end of the mail.</p>

<p><b>Time & Date : 22<sup>nd</sup> May ,2020; 20:00 IST</b></p>

<p>Program Schedule:</p>

<p>Participant Login: 19:55 IST Speaker <br>
Introduction & Overview : Dr Sameer SM, Chair - Strategic Initiative for Academic Engagements, IEEE Kerala Section.<br> 
Webinar and Q&A : Dr SK Ramesh.</p>

<p>Program Protocol:</p>

<p>- Participants are requested to login by 19:55 IST.<br> 
- Participants make sure that their microphone always remains muted.<br> 
- All participants, except for speakers, are requested to turn off their videos so as to conserve bandwidth. <br>
- Participants can type in their questions in the chat window anytime during the program. Selected questions will be read out by the moderator.<br>
 - Participants can also type in their queries in the chat window in case of any technical issues.<br></p>

<p><b>Participants will be awarded an e-certificate on completion of the webinar.</b> However delegates need to attend the session from beginning to end and provide the feedback using the link given below after the webinar but latest by 23<sup>rd</sup> May, 2020.<br>
https://bit.ly/AccrWSFeedback</p>
<p style="color:blue">Meeting Link:<br>
Join Zoom Meeting:<br>
https://us02web.zoom.us/j/88319014148?pwd=VXJQR0h1OUd0ajMwdEtaS2ZzT3d2QT09 <br>
<br>
ID: 883 1901 4148<br>
Password: 123456<br>
<br>
Find your local number: https://us02web.zoom.us/u/kbm5pGYYH4</p>



<p>Looking forward to your presence during the Webinar.</p>

<p>Regards</p>

<p style="color:grey">Dr Sameer SM<br>
Chair - Strategic Initiative for Academic Engagement | IEEE KS.<br>
Professor - Dept. of ECE | National Institute of Technology, Calicut.<br>
Head - Centre for Training & Placement | National Institute of Technology, Calicut.<br></p>

    </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )
        