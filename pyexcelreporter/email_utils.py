#!/usr/bin/env python
# coding: utf-8

# This little project is hosted at: <https://gist.github.com/1455741>
# Copyright 2011-2012 Álvaro Justen [alvarojusten at gmail dot com]
# Modified 2018 by Nathan Morin https://github.com/nathanamorin
# License: GPL <http://www.gnu.org/copyleft/gpl.html>

import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from mimetypes import guess_type
from email.encoders import encode_base64
from smtplib import SMTP


def get_email(email):
    if '<' in email:
        data = email.split('<')
        email = data[1].split('>')[0].strip()
    return email.strip()

class Email(object):
    def __init__(self, from_email, to_emails, subject, message, message_type='plain',
                 attachments=None, cc=None, message_encoding='us-ascii'):
        self.email = MIMEMultipart()
        self.email['From'] = from_email
        self.email['To'] = ", ".join(to_emails) if type(to_emails) is list else to_emails
        self.email['Subject'] = subject
        if cc is not None:
            self.email['Cc'] = cc
        text = MIMEText(message, message_type, message_encoding)
        self.email.attach(text)
        if attachments is not None:
            for filename in attachments:
                mimetype, encoding = guess_type(filename)
                mimetype = mimetype.split('/', 1)
                fp = open(filename, 'rb')
                attachment = MIMEBase(mimetype[0], mimetype[1])
                attachment.set_payload(fp.read())
                fp.close()
                encode_base64(attachment)
                attachment.add_header('Content-Disposition', 'attachment',
                                      filename=os.path.basename(filename))
                self.email.attach(attachment)

    def __str__(self):
        return self.email.as_string()


class EmailConnection(object):
    def __init__(self, server, port, username, password, secure=True):
        self.server = server
        self.port = port
        self.username = username
        self.password = password
        self.secure = secure
        self.connect()

    def connect(self):
        self.connection = SMTP(self.server, self.port)
        self.connection.ehlo()
        if self.secure:
            self.connection.starttls()
        self.connection.ehlo()
        if self.username and self.password:
            self.connection.login(self.username, self.password)

    def send(self, message):

        from_email = message.email['From']
        if 'Cc' not in message.email:
            cc_emails = []
        else:
            cc_emails = message.email['Cc'].split(',')
        all_emails = message.email['To'].split(',') + cc_emails
        all_emails = [get_email(complete_email) for complete_email in all_emails if all_emails != '' and all_emails is not None]
        message = str(message)
        return self.connection.sendmail(from_email, all_emails, message)

    def close(self):
        self.connection.close()