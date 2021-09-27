import smtplib, ssl, io

from docxtpl import DocxTemplate, RichText
from flask import send_file, request

from email.mime.base import MIMEBase 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText
from email import encoders


fromaddr = 'pochtastar@gmail.com'
toaddr = 'ilya.gutsko@kaspersky.com'

msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = 'This is the subject of my email'
body = 'This is the body of my email'

msg.attach(MIMEText(body))

context = { 'company_name' : "World company" }

doc = DocxTemplate(r"C:\Users\gutsko_i\Desktop\portal\test\template.docx")
doc.render(context)
file_stream = io.BytesIO()
doc.save(file_stream)
file_stream.seek(0)

doc1 = send_file(file_stream, as_attachment=True, attachment_filename='msg.docx')

files = [doc1]


for filename in files:

    attachment = open(filename, 'rb')
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)

    part.add_header("Content-Disposition",
    f"attachment; filename= {filename}")

    msg.attach(part)

msg = msg.as_string()

server = smtplib.SMTP('smtp.gmail.com:587', timeout=120)
server.ehlo()
server.starttls()
server.ehlo()
server.login(fromaddr, '')
server.sendmail(fromaddr, toaddr, msg)
server.quit()
