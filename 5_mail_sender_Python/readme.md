
## Python Mail Sender
<img src="./pictures/photo_2024-08-06_10-01-01.jpg" width=300 align="left">


If you have this kind of task, Python allows sending email messages using builtin libraries (SMTP, email, IMAP).  

To make the work with mailing list easier without deep immersion into these libraries I wrote a [small module](mail_sender.py) `email_sender`.  

The central object in this module is the class `EmailSender` and it takes 4 parameters for its initialization: smtp_server, smtp_port, email_from, password. You can get these parameters from any free email box (mail.ru, gmail.com).  

This class has some methods:
* `send`  - method which allows us to send 1 message to many target emails. The message can contain 1 or many attachments.
* `send_messages` - if we need to send several messages to many emails: each message to all emails, we should use this method and specify lists of message texts and subjects as the corresponding parameters.
* `send_files` - if a message text is unchanged and an attachment is only being changed, you should use the method send files. You need to use lists of lists to specify many attachments for each letter.
* `send_by_table` - method which takes only dict as a single parameter. You can specify each sending separatelly by using a special table or dict. This method use every row of the given table step-by-step to define parameters of sending.

See the full [Jupyter notebook](./mail_sender1_1.ipynb) for details.
