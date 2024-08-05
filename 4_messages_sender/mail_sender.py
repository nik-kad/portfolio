#!/usr/bin/env python
# coding: utf-8

from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formatdate
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import imaplib
import smtplib
import os
import time
import base64
import sys
import itertools
import ast

class EmailSender:
    def __init__(self, smtp_server: str,
                 email_from: str,
                 password: str,
                 smtp_port: int = 25,
                 imap_server: str | None = None,
                 imap_port: int | None = None):

        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        
        self.imap_server = imap_server
        self.imap_port = imap_port

        self.email_login = email_from
        self.password = password
        try:
            smtp = smtplib.SMTP(self.smtp_server, self.smtp_port)
            
            print('SMTP - ok')
            smtp.quit()
        except:
            print('SMTP - error. Check server name and port number or internet connection')
            sys.exit()

        if imap_server:
            try:
                imap = imaplib.IMAP4(imap_server, imap_port)
                imap.starttls()
                self.imap = imap
                print('IMAP - ok')
            except:
                print('IMAP - error. Check server name and port number.')
                print('Messages synchronization is unavailable! Only sending')

    def send(self, emails_to: list | str, subject: str = '', message_text: str = '',
             attachment_paths: str | list | None = None, email_from: str = None):

        if not email_from:
            email_from = self.email_login

        if isinstance(emails_to, str):
            emails_to = [emails_to]
        elif isinstance(emails_to, (list, tuple)):
            pass
        else:
            print('Wrong `emails_to` format. Use list or str')
            sys.exit()

        print('Creating letter')
        # create letter
        message = MIMEMultipart()
        message['From'] = email_from
        
        message['Subject'] = Header(str(subject), 'utf-8')
        message['Date'] = formatdate(localtime=True)
        message.attach(MIMEText(str(message_text), 'html', 'utf-8'))
        print('Ok')

        
        if attachment_paths:
            print('Adding attachments')
            # adding attachment
            if isinstance(attachment_paths, str):
                attachment_paths = [attachment_paths]
                
            elif isinstance(attachment_paths, (list, tuple)):
                pass
            else:
                print('Wrong type of the attacments_paths. Use list of str or str')
            
            for path in attachment_paths:
                try:
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(open(path, 'rb').read())
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(path)}"')
                    message.attach(attachment)
                    print(f'Path: {path} has been added successfully')
                except:
                    print(f'Can`t add the attachment from the path: {path}. Check the path.')
        
        print('Establishing connection')
        smtp = smtplib.SMTP(self.smtp_server, self.smtp_port)
        print(smtp.starttls())
        print(smtp.ehlo())
        print(smtp.login(self.email_login, self.password))
        print('Connection - Ok')
        

        for email in emails_to:

            if not 'To' in message:
                message['To'] = email
            else:
                message.replace_header('To', email)
            print(f'Sending to {email}')
            smtp.sendmail(self.email_login, email, message.as_string())
            print('Ok')
            
        smtp.quit()

    def send_messages(self, messages_list: list, emails_to: list | str, subject: str | list = '', 
                      attachment_paths: str | list | None = None, email_from: str = None):
        if not isinstance(messages_list, (list, tuple)):
            print('Wrong `messages_list` format. Use list')
            sys.exit()

        if isinstance(subject, list):
            msglist_len = len(messages_list)
            if len(subject) > msglist_len:
                subject = subject[:msglist_len]
            for i, msg, sbj in itertools.zip_longest(range(1, msglist_len+1), messages_list, subject, fillvalue=''):
                print(f'Sending message {i} of {msglist_len}')
                self.send(emails_to=emails_to, subject=sbj, message_text=msg,
                          attachment_paths=attachment_paths, email_from=email_from)
                print('----------------------------------------------------------------------------------------------------------------')
                
        else:
            for i, msg in enumerate(messages_list):
                print(f'Sending message {i+1} of {msglist_len}')
                self.send(emails_to=emails_to, subject=subject, message_text=msg,
                          attachment_paths=attachment_paths, email_from=email_from)
                print('----------------------------------------------------------------------------------------------------------------')
        print('Done')

    def send_files(self, attachments_paths: list, emails_to: list | str, subject: str = '', 
                      message_text: str = '', email_from: str = None):
        
        if not isinstance(attachments_paths, (list, tuple)):
            print('Wrong `attachments_paths` format. Use list')
            sys.exit()

        atchpath_len = len(attachments_paths)
        for i, atch in enumerate(attachments_paths):
            print(f'Sending attachments {i+1} of {atchpath_len}')
            
            self.send(emails_to=emails_to, subject=subject, message_text=message_text,
                      attachment_paths=atch, email_from=email_from)
            print('----------------------------------------------------------------------------------------------------------------')
        print('Done')

    def send_by_table(self, data_for_sending: dict):
        if not isinstance(data_for_sending, dict):
            print('Wrong `data_for_sending` format. Use dict')
            sys.exit()

        if not 'emails_to' in data_for_sending:
            print('No emails to send!. Use `emails_to` key for dict')
            sys.exit()
            
        if not 'messages' in data_for_sending:
            data_for_sending['messages'] = ['']
            
        if not 'subjects' in data_for_sending:
            data_for_sending['subjects'] = ['']
            
        if not data_for_sending['emails_to']:
            print('The list of target emails is empty!. You should specify even one email')
            sys.exit()

        emaillist_len = len(data_for_sending['emails_to'])
        for param in data_for_sending:
            if len(data_for_sending[param]) > emaillist_len:
                data_for_sending[param] = data_for_sending[param][:emaillist_len]
                
        if not 'attachments_paths' in data_for_sending and not 'emails_from' in data_for_sending:
            for i, eml, msg, sbj in itertools.zip_longest(range(1, emaillist_len+1),  data_for_sending['emails_to'],
                                                          data_for_sending['messages'], data_for_sending['subjects'], fillvalue=''):
                print(f'Sending message {i} of {emaillist_len}')
                try:
                    eml = ast.literal_eval(eml)
                except:
                    pass
                    
                self.send(emails_to=eml, subject=str(sbj), message_text=str(msg),
                          attachment_paths=None, email_from=None)
                print('---------------------------------------------------------------------------')
        
        elif not 'attachments_paths' in data_for_sending:
            for i, eml, msg, sbj, emlf in itertools.zip_longest(range(1, emaillist_len+1),  data_for_sending['emails_to'],
                                                          data_for_sending['messages'], data_for_sending['subjects'],
                                                                data_for_sending['emails_from'], fillvalue=''):
                
                print(f'Sending message {i} of {emaillist_len}')
                try:
                    eml = ast.literal_eval(eml)
                except:
                    pass
                self.send(emails_to=eml, subject=str(sbj), message_text=str(msg),
                          attachment_paths=None, email_from=emlf)
                print('---------------------------------------------------------------------------')
            
        elif not 'emails_from' in data_for_sending:
            for i, eml, msg, sbj, atch in itertools.zip_longest(range(1, emaillist_len+1),  data_for_sending['emails_to'],
                                                          data_for_sending['messages'], data_for_sending['subjects'],
                                                                data_for_sending['attachments_paths'], fillvalue=''):
                print(f'Sending message {i} of {emaillist_len}')
                try:
                    eml = ast.literal_eval(eml)
                except:
                    pass
                try:
                    atch = ast.literal_eval(atch)
                except:
                    pass
                    
                self.send(emails_to=eml, subject=str(sbj), message_text=str(msg),
                          attachment_paths=atch, email_from=None)
                print('---------------------------------------------------------------------------')
        else:
            for i, eml, msg, sbj, atch, emlf in itertools.zip_longest(range(1, emaillist_len+1),  data_for_sending['emails_to'],
                                                          data_for_sending['messages'], data_for_sending['subjects'],
                                                                data_for_sending['attachments_paths'], data_for_sending['emails_from'],
                                                                      fillvalue=''):
                print(f'Sending message {i} of {emaillist_len}')
                try:
                    eml = ast.literal_eval(eml)
                except:
                    pass
                try:
                    atch = ast.literal_eval(atch)
                except:
                    pass
                self.send(emails_to=eml, subject=str(sbj), message_text=str(msg),
                          attachment_paths=atch, email_from=emlf)
                print('---------------------------------------------------------------------------')
        print('Done')
 
