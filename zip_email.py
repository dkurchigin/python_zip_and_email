import zipfile
import re
import os
import smtplib
import json
import time
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart


directory = '.'
files = os.listdir(directory)
inputfile_pattern = '.*\.(xlsx|xls)$'


def load_parameters(file_with_parameters='settings.json'):
    with open(file_with_parameters, "r", encoding='utf-8') as read_file:
        parsed_json = json.load(read_file)

        mail_dict = {
            'server': parsed_json["server"],
            'passwd': parsed_json["passwd"],
            'from': parsed_json["from"],
            'to': parsed_json["to"],
            'subject': parsed_json["subject"]
        }
        mail_dict['message'] = 'From: {}\nTo: {}\nSubject: {}\n\ntest'.format(
            mail_dict['from'], mail_dict['to'], mail_dict['subject'])
        return mail_dict


def try_send_email(mail_dict, attachment_file):
    smtp_obj = smtplib.SMTP_SSL(mail_dict['server'])
    smtp_obj.set_debuglevel(1)
    smtp_obj.login(mail_dict['from'], mail_dict['passwd'])

    msg = MIMEMultipart()
    msg['Subject'] = mail_dict['subject']
    msg['From'] = mail_dict['from']
    msg['To'] = mail_dict['to']
    msg['Date'] = time.strftime('%a, %d %b %Y %H:%M:%S %z')

    txt = MIMEText("test mail", 'plain', 'utf-8')
    msg.attach(txt)

    att = MIMEApplication(open(attachment_file, 'rb').read(), _subtype='zip')
    att["Content-Disposition"] = 'attachment; \n\tfilename="{}"'.format(attachment_file)
    msg.attach(att)

    smtp_obj.sendmail(mail_dict['from'], mail_dict['to'], msg.as_string())
    smtp_obj.quit()


def zip_zip(xlsx_file):
    try:
        zip_file_pattern = re.sub(r'\.(xlsx|xls)$', '.zip', xlsx_file)
        zipped_file = zipfile.ZipFile(zip_file_pattern, 'w')
        zipped_file.write(xlsx_file)
        print("Архив {} создан".format(zip_file_pattern))
        zipped_file.close()
        return zip_file_pattern
    except:
        print("Ошибка создания архива!")


for file in files:
    if re.match(inputfile_pattern, file):
        zipped_file = zip_zip(file)
        loaded_parameters = load_parameters()
        try_send_email(loaded_parameters, zipped_file)
