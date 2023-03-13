
import csv
import os,re
import datetime
import calendar
import pandas as pd
from datetime import date
from exchangelib import Configuration,Account, HTMLBody,OAuth2Credentials, OAUTH2, IMPERSONATION, Build, Version, UTC_NOW, Message, Mailbox
from exchangelib import Message, Mailbox, FileAttachment
from email.mime.image import MIMEImage
from email.mime.text import MIMEText


#Setting email credentials to connect to the email
username = 'support-emails@system.modata.com'
credentials = OAuth2Credentials(client_id='8df56b54-c349-4a7b-bb41-a3ac352ec879', client_secret='9Ca8Q~QXx..eWRsRT6.zKE2bzYALoFNnKXIh1a86', tenant_id='23349441-f82d-4441-bcf0-850c5e2d83f8')
version = Version(build=Build(15, 0, 12, 34))

#connection to the outlook server and login using the credentials
config = Configuration(service_endpoint = 'https://outlook.office365.com/EWS/Exchange.asmx',credentials=credentials,version=version,auth_type=OAUTH2)
account = Account(username, credentials=credentials, autodiscover=False, config=config, access_type=IMPERSONATION)

run_date = "2023-02-28"
employee_status_data = list()
output_array = [["employee name & surname","project name", "client name","task name","total logged time"]]
    
def convert_minutes_to_decimals(time_string):
    #convert minutes to decimals, matching the time string to the regex
    regex = re.compile("[0-9][0-9]*:[0-5][0-9]")
    if regex.match(time_string) is not None:
        time_string_array_split = time_string.split(':')
        minutes_to_decimals = int(time_string_array_split[1]) / 60

        return minutes_to_decimals 

    else: 

        error_msg = "Failed to convert time string. Incorrect format [{}]".format(time_string)
        raise Exception(error_msg)

def get_zoho_projects(name_surname):
    #reading monthtimesheets file 
    monthtimesheets_file_path = "./input_files/{}/MonthTimesheetReport.csv".format(run_date)
    with open(monthtimesheets_file_path) as csv_file:
        csv_reader = csv.reader(csv_file,quotechar='"', delimiter=',')
        line_count = 0
        for row in csv_reader:
            csv_file_array = []
            if line_count == 0:
                line_count += 1
                continue
            if row[0].strip().lower() == "total":
                break
        
            row_name = "{}".format(row[1])
            if name_surname.strip().lower() == row_name.strip().lower():
                #csv_file_row = "{},{},{},\"{}\",{}".format(row[1].strip(),row[4].strip(),row[3].strip(),row[5].strip(),row[-1].strip())
                #convert the time string minutes to decimal
                minutes_in_decimals = convert_minutes_to_decimals(str(row[-1].strip()))
                time_string_splited = row[-1].strip().split(":")
                converted_time = int(time_string_splited[0]) + minutes_in_decimals
                #round off the time to the nearest 1
                converted_time = "{:.2f}".format(converted_time)
                #add the data into the output array
                csv_file_array.append(row[1].strip())
                csv_file_array.append(row[4].strip())
                csv_file_array.append(row[3].strip())
                csv_file_array.append(row[5].strip())
                csv_file_array.append(converted_time)
                output_array.append(csv_file_array)
            
            line_count += 1             

def get_zoho_tickets(name_surname):
    #Set the directory path to the folder path.check if the file name start with ExportReport_
    folder_path = "./input_files/{}/".format(run_date)
    file_name = ""
    for f in os.listdir(folder_path):
        if re.match('ExportReport_', f):
            file_name = f
            break
        
    #Reading the export report ticket file 
    exportreport_file_path = "./input_files/{}/{}".format(run_date,file_name)
    with open(exportreport_file_path) as csv_file:
        csv_reader = csv.reader(csv_file,quotechar='"', delimiter=',')
        line_count = 0
        
        #looping through the data in csv_reader line by line adding that to row
        for row in csv_reader:
            csv_file_array = []
            if line_count < 5:
                line_count += 1
                continue
            
            if line_count > 4:
                #check if the file has spaces, continue
                if len(row) == 0:
                    continue
                
                #check if the row has spaces, continue
                if not row[0]:
                    continue
                
                #check for the word "total records" in a string and if teh word is found, break 
                if row[0].lower().find("total records") != -1:
                    break
                    
                row_name = "{}".format(row[0])

                if name_surname.strip().lower() == row_name.strip().lower():
                    #Converting the total time logged minutes to seconds using pandas library
                    ts = pd.to_datetime(row[-1].strip())
                    ts_rounded = ts.round(freq='T')
                    ts_rounded = ts_rounded.strftime("%H:%M")
                    #converting  the time string minutes to decimals and adding that to the hours
                    minutes_in_decimals = convert_minutes_to_decimals(ts_rounded)
                    time_string_splited = row[-1].strip().split(":")
                    converted_time = int(time_string_splited[0]) + minutes_in_decimals
                    #rounding oof the time to the nearest 1
                    converted_time = "{:.2f}".format(converted_time)
                    
                    csv_file_array.append(row[0].strip())
                    csv_file_array.append("Support")
                    csv_file_array.append(row[1].strip())
                    csv_file_array.append(row[3].strip())
                    csv_file_array.append(converted_time)
                    output_array.append(csv_file_array)
                    
        line_count += 1
    
def get_paid_booked_leave(employee_details):
    leave_book_file_path = "./input_files/{}/Leave_booked_and_balance.csv".format(run_date)
    with open(leave_book_file_path) as csv_file:
        csv_reader = csv.reader(csv_file,quotechar='"', delimiter=',')
        line_count = 0
        
        #looping through the data in csv_reader line by line adding that to row
        for row in csv_reader:
            csv_file_array = []
            #check if line count is less than 2 to skip the headers
            if line_count < 2:
                line_count += 1
                continue
    
            row_name = "{}".format(row[1])
            
            if employee_details[0].strip().lower() == row_name.strip().lower() and employee_details[1].strip().lower() == "permanent":
                convert_hours_to_string = float(row[11].strip()) * 8
                convert_hours_to_string = str(convert_hours_to_string).replace(".0", ":00")
                #converting  the time string minutes to decimals and adding that to the hour
                minutes_in_decimals = convert_minutes_to_decimals(convert_hours_to_string)
                time_string_splited = convert_hours_to_string.split(":")
                converted_time = int(time_string_splited[0]) + minutes_in_decimals
                #round off the time to the nearest 1
                converted_time = "{:.2f}".format(converted_time)
                
                csv_file_array.append(row[1].strip())
                csv_file_array.append("Paid Annual Leave")
                csv_file_array.append("Paid Annual Leave")
                csv_file_array.append("Paid Annual Leave")
                csv_file_array.append(converted_time)
                output_array.append(csv_file_array)
                

        line_count += 1
        
def get_unpaid_booked_leave(employee_details):
    leave_book_file_path = "./input_files/{}/Leave_booked_and_balance.csv".format(run_date)
    with open(leave_book_file_path) as csv_file:
        csv_reader = csv.reader(csv_file,quotechar='"', delimiter=',')
        line_count = 0
        
        #looping through the data in csv_reader line by line adding that to row
        for row in csv_reader:
            csv_file_array = []
            #check if line count is less than 2 to skip the headers
            if line_count < 2:
                line_count += 1
                continue
        
            row_name = "{}".format(row[1])
            
            if employee_details[0].strip().lower() == row_name.strip().lower() and employee_details[1].strip().lower() == "permanent":
                convert_hours_to_string = float(row[19].strip()) * 8
                convert_hours_to_string = str(convert_hours_to_string).replace(".0", ":00")
                #converting  the time string minutes to decimals and adding that to the hour
                minutes_in_decimals = convert_minutes_to_decimals(convert_hours_to_string)
                time_string_splited = convert_hours_to_string.split(":")
                converted_time = int(time_string_splited[0]) + minutes_in_decimals
                #round off the time to the nearest 1
                converted_time = "{:.2f}".format(converted_time)
                
                csv_file_array.append(row[1].strip())
                csv_file_array.append("Unpaid Annual Leave")
                csv_file_array.append("Unpaid Annual Leave")
                csv_file_array.append("Unpaid Annual Leave")
                csv_file_array.append(converted_time)
                output_array.append(csv_file_array)
                

        line_count += 1
            
def get_public_holidays():
    public_holiday_array = {}
    #Getting first 3 letters of the month name from the run date variable 
    date_object = datetime.datetime.strptime(run_date, '%Y-%m-%d')
    month_name = calendar.month_name[date_object.month]
    #Taking only the first 3 letters of the month name(eg Apr)
    month_name = month_name[:3]
 
    #reading the holidays calendar
    holidays_file_path = "./lookup_files/calendars/holidays_2023.csv"
    with open(holidays_file_path) as csv_file:
        csv_reader = csv.reader(csv_file,quotechar='"', delimiter=',')
        line_count = 0
        
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
                continue
            
            #check if the row has spaces and continue
            if len(row) == 0 or len(row) == 1:
                continue
            
            #check if the run time month name is the same as the name in the calendar , then add hours if you have public holidays
            if row[1][:3].lower() == month_name.lower():
                public_holiday_array[row[0].strip()] = "8.0"
    
                
    return public_holiday_array

    #check if we have a file with today's date

#get the output file with today's date
def get_output_file():
    folder_path = "./output_files"
    todays_date = date.today()
    todays_date = todays_date.strftime("%Y%m%d")

    try:
        for folder in os.listdir(folder_path):
            get_date_in_file = folder.split('-')
            if re.match(todays_date , get_date_in_file[0]):
                return folder
    
    except Exception as error:
        print("No files find today", error)
        exit()
        
    return None

#send email with the output file atatched
def send_email(account,subject,body,recipients, attachments=None):
    to_recipients = []
    for recipient in recipients:
        to_recipients.append(Mailbox(email_address=recipient))
        
    # with open('./email_template/template.html') as f:contents = f.read()
    # contents = re.sub('!!content!!', 'Kind regards, \n\n Abigail Hlongwani', contents)
 
    # Create message
    m = Message(account=account,
                folder=account.sent,
                subject=subject,
                body=body,
                # body=HTMLBody(contents),
                to_recipients=to_recipients,
                )
    
    logoname = './img/MoData.png'
    with open(logoname, 'rb') as fp:
        logoimg = FileAttachment(name=logoname, content=fp.read())
    m.attach(logoimg)

    for attachment_name, attachment_content in attachments or []:
        file = FileAttachment(name=attachment_name, content=attachment_content)
        m.attach(file)
    m.send_and_save()
    
result = get_output_file()
attachments = []

if result != None:
    with open('./output_files/{}'.format(result), 'rb') as f:
        content = f.read()
    attachments.append((result, content))
else:
    print("No file found")
        
def main():
    #loading employee status file info into array
    with open('./lookup_files/employee_status/employee_status.csv', newline='') as csvfile:
        employee_status_data = list(csv.reader(csvfile))
    
    line_count = 0
    
    #get the public holidays 
    public_holidays = get_public_holidays()
    
    for each_employee in employee_status_data:
        if line_count == 0:
            line_count += 1
            continue
        #check public holidays and if the employee is permanent, then add the public holidays to the output file
        if bool(public_holidays) == True and each_employee[1].strip().lower() == "permanent":
            for key, value in public_holidays.items():
                output_array.append([each_employee[0],"Public Holiday","Public Holiday", key,value])

        get_zoho_projects(each_employee[0])
        get_zoho_tickets(each_employee[0])
        get_paid_booked_leave(each_employee)
        get_unpaid_booked_leave(each_employee)
 
        line_count += 1

    #writing output_array into an output file
    now = datetime.datetime.now()
    date = now.strftime("%Y%m%d")
    file_name = "{}-{}-ZohoTimeSheets.csv".format(date,run_date.replace("-",""))

    full_file_name = './output_files/'+file_name  
    try:     
        
        with open(full_file_name, 'w', newline='') as file:
            #writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)     
            writer = csv.writer(file)
            writer.writerows(output_array)
            #writer.writerows([x.split(',') for x in output_array])
    except Exception as error:
        print("Failed to write file", error)
        
    send_email(account, 'Zoho Monthly Timesheet Output File','Please Find attached Timesheet output file.\nPlease let me know if there is anything missing.\n\n King regards,\nAbigail Hlongwani', ['abigail@modata.com'],attachments=attachments)

main()


