# import pandas as pd

# # read the Excel file
# df = pd.read_excel('E:\Abid_Data\p\email-database-mustafa(1) (1).xlsx')

# # get the first column
# first_column = df.iloc[:, 0]
# c=1
# for i in first_column:
#     print(f"{c}   :{i}")
#     c+=1
   





import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd

# # read the Excel file
# df = pd.read_excel('E:\Abid_Data\p\email-database-mustafa(1) (1).xlsx')

# # get the first column
# first_column = df.iloc[:, 0]
c=1
for i in []:
    print(f"{c}   :{i}")
    c+=1
    sender ='YourEmail@gmail.com' 
    password = 'YourPassword'   

    recipient = i
    subject = 'New on the Business History channel on YouTube'
    body =  """\
            <html>

    <head>
    </head>

    <body>
        <div
            style="max-width:655px; font-family:Georgia, 'Times New Roman', Times, serif; font-size:13.5pt; line-height:1.6;">
            <p style="max-width:655px;">Hi!</p>
            <p style="max-width:655px;">Our latest video is up on YouTube.</p>
            <p style="max-width:655px;"><strong>Vance Packard: The man who exposed advertising hype and planned
                    obsolescence</strong></p>
            <p style="max-width:655px;">See it here: <a
                    href="https://www.youtube.com/@BusinessHistory?sub_confirmation=1&utm_source=newsletter&utm_medium=email&utm_campaign=YouTube">https://www.youtube.com/watch?v=xYy3Kf0kw30</a>
            </p>
            <p style="max-width:655px;">
                In the 1950s Vance Packard exposed the advertising hype and planned obsolescence encouraged by companies,
                which were forever trying to sell stuff to their customers, whether they really needed it or not. (They are
                doing it even more aggressively today.)
                See this video and many more on the Business History channel on YouTube. Videos on the history of industry,
                companies, technology, people, and more. Here:
                <a
                    href="https://www.youtube.com/@BusinessHistory?sub_confirmation=1&utm_source=newsletter&utm_medium=email&utm_campaign=YouTube">https://www.youtube.com/@BusinessHistory</a>
            </p>
            <p style="max-width:655px;">
                <a
                    href="https://www.youtube.com/@BusinessHistory?sub_confirmation=1&utm_source=newsletter&utm_medium=email&utm_campaign=YouTube"><img
                        style="max-width:655px;"
                        src="https://businesshistory.domain-b.com/images/business-history-channel.jpg?utm_source=newsletter&utm_medium=email&utm_campaign=YouTube" /></a>
            </p>
            <p style="max-width:655px;">Please also forward this mail to friends, or share the video link with them.</p>
            <p style="max-width:655px;">We hope you have already subscribed to the Business History Channel. If you haven’t,
                you should, for then you will be alerted every time we upload a new video on YouTube.</p>
            <p style="max-width:655px;">Do also let us know what you think of the videos. You can write your comments in the
                space below the videos.</p>
            <p style="max-width:655px;">
                Best wishes!
                <br />The Business History Team
            </p>
        </div>
    </body>
    </htm>
            """


    # create message object instance
    message = MIMEMultipart('alternative')
    message['From'] = sender
    message['To'] = recipient
    message['Subject'] = subject

    # create HTML part of message
    html_part = MIMEText(body, 'html')

    # attach HTML part to message
    message.attach(html_part)

    # create SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # login to sender email account
    s.login(sender, password)

    # send the message via the SMTP server
    s.send_message(message)

    # terminate the SMTP session
    s.quit()

    print("HTML email sent successfully.")
print("Sab Send Ho Gya Hai ....!")
