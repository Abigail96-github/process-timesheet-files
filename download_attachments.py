
from exchangelib import Configuration,Account,OAuth2Credentials, OAUTH2, IMPERSONATION, Build, Version, UTC_NOW, Message, Mailbox
from datetime import datetime, date, timedelta
import os,re

#Setting email credentials to connect to the email
username = 'zoho.reporting@system.modata.com'
credentials = OAuth2Credentials(client_id='8df56b54-c349-4a7b-bb41-a3ac352ec879', client_secret='9Ca8Q~QXx..eWRsRT6.zKE2bzYALoFNnKXIh1a86', tenant_id='23349441-f82d-4441-bcf0-850c5e2d83f8')
version = Version(build=Build(15, 0, 12, 34))

#connection to the outlook server and login using the credentials
config = Configuration(service_endpoint = 'https://outlook.office365.com/EWS/Exchange.asmx',credentials=credentials,version=version,auth_type=OAUTH2)
account = Account(username, credentials=credentials, autodiscover=False, config=config, access_type=IMPERSONATION)

folders= account.inbox.all()
path = ""   
def get_monthly_project_attachment():
    for item in account.inbox.filter(subject__exact='Monthly Project Timesheets').order_by('-datetime_received')[:1]:
        for attachment in item.attachments:
            today = date.today()
            todays_month = today.strftime("%m")
            prev_month = today.replace(day=1) - timedelta(days=1)
            path = ""
        
        #check which month in todays date and create a folder with the prev month date
            if int(todays_month) == 1:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                            
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)

            elif int(todays_month) == 2:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 3:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()

                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 4:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 5:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 6:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
            elif int(todays_month) == 7:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 8:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 9:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 10:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 11:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit() 
                    
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            elif int(todays_month) == 12:
                path = './input_files/{}'.format(prev_month)
                try: 
                    os.mkdir(path) 
                except OSError as error: 
                    print(error)
                    exit()
                            
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)
            else: 
                print('invalid month')
                
def get_tickets_attachment():
    for item in account.inbox.filter(subject__exact='Ticket Timesheets').order_by('-datetime_received')[:1]:
        for attachment in item.attachments:

            today = date.today()
            prev_month = today.replace(day=1) - timedelta(days=1)
            path = './input_files/{}'.format(prev_month)
            
            if os.path.exists(path) == True:
                with open(path+'/{}'.format(attachment.name), 'wb') as f:
                    f.write(attachment.content)              

get_monthly_project_attachment()
get_tickets_attachment()


# from process_timesheets import main 
# main()


# import email
# import base64
# import os
# import imaplib

# #access email details
# email_user = "zoho.reporting@system.modata.com"
# email_pass = "Von65606"

# #connecting to the email
# mail = imaplib.IMAP4_SSL("system.modata.com",465)
# print('here')
# #logging into the email
# mail.login(email_user, email_pass)
# #selecting folder we want to read from(Inbox)
# mail.select('Inbox')

# #search the subject you want to read from and the latest one
# #type, data = mail.search('search', none, "Monthly Project Timesheets")
# type, data = mail.search(None, 'ALL')
# mail_ids = data[0]
# id_list = mail_ids.split()

# #RFC822 defines an electronic message format consisting of header fields and an optional message body
# for num in data[0].split():
#     typ, data = mail.fetch(num, '(RFC822)' )
#     raw_email = data[0][1]
#     # converts byte literal to string removing b''
#     raw_email_string = raw_email.decode('utf-8')
#     email_message = email.message_from_string(raw_email_string)
#     # downloading attachments
#     #walk() generates the file names in a directory tree by walking the tree either top-down or bottom-up
#     for part in email_message.walk():
#         # this part comes from the snipped I don't understand yet... 
#         if part.get_content_maintype() == 'multipart':
#             continue
#         if part.get('Content-Disposition') is None:
#             continue
#         fileName = part.get_filename()
#         if bool(fileName):
#             filePath = os.path.join('.input_files', fileName)
#             if not os.path.isfile(filePath) :
#                 fp = open(filePath, 'wb')
#                 fp.write(part.get_payload(decode=True))
#                 fp.close()
#             subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
#             print('Downloaded "{file}" from email titled "{subject}" with UID {uid}.'.format(file=fileName, subject=subject, uid=latest_email_uid.decode('utf-8')))