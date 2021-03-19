
import imghdr
import smtplib
from email.message import EmailMessage

My_Email = 'snb.khurram@gmail.com'


def send_email(to):
    message = EmailMessage()
    message['subject'] = "Business Intro"
    message['from'] = My_Email
    message['to'] = to
    message.set_content('Hello this is bulk email tester')

    html_message = open('Html Demo.html').read()
    message.add_alternative(html_message, subtype='html')

    # with open('logo.png', 'rb') as attach_file:
    #     image_name = attach_file.name
    #     image_type = imghdr.what(attach_file.name)
    #     image_data = attach_file.read()
    # message.add_attachment(image_data, maintype="image", subtype=image_type, filename=image_name)

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login('snb.khurram@gmail.com', 'khurramsaeed03360736736')
        smtp.send_message(message)


send_email(My_Email)
print('All Emails Sent Successfully')
