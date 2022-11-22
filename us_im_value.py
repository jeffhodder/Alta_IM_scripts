import imaplib, email
import os
import csv
from openpyxl import load_workbook
from datetime import datetime, timedelta,date
from os.path import exists
import pyodbc

yesterday = datetime.today() - timedelta(days=1)
raw_date = yesterday.strftime('%m%d%y')



unwanted_days = [5, 6]
wanted_days = [0, 1, 2, 3, 4]
if yesterday.weekday() in unwanted_days:
    if datetime.today().weekday() == 0:
        last_friday = datetime.today() - timedelta(days=3)
        raw_date_friday = last_friday.strftime('%m%d%y')
        today_date = str(raw_date_friday[4:7] + raw_date_friday[0:2] + raw_date_friday[2:4])
        asof_day = int(raw_date_friday[2:4])
        asof_month = int(raw_date_friday[0:2])
        asofyear = raw_date_friday[4:7]
    else:
        quit()
else:
    today_date = str(raw_date[4:7] + raw_date[0:2] + raw_date[2:4])
    asof_day = int(raw_date[2:4])
    asof_month = int(raw_date[0:2])
    asofyear = raw_date[4:7]

# uncomment below lines for retrospective dates
# today_date = '221012'
# asof_day = 12
# asof_month = 10
# asofyear = '22'
file_location = str('\\\\altfps\\arcadiagroup$\Midoffice\IM Attachments\Houston IM Files\ARCUS_CONFORM_20' + today_date + '.csv')


# if exists(file_location) is True:
#     print('File already in Houston IM folder. To run the script comment out lines 38-40')
#     quit()

# user = 'j.adjaho@altatrading.com'
# password = 'table20!longerchair'
# imap_url = 'imap-mail.outlook.com'
attachment_dir = '\\\\altfps\\arcadiagroup$\Midoffice\IM Attachments\Houston IM Files'

# def get_attachments(msg):
#     for part in msg.walk():
#         if part.get_content_maintype == 'multipart':
#             continue
#         if part.get('Content-Disposition') is None:
#             continue
#         fileName = part.get_filename()

#         if bool(fileName):
#             filePath = os.path.join(attachment_dir, fileName)
#             with open(filePath, 'wb') as f:
#                 f.write(part.get_payload(decode=True))

# con = imaplib.IMAP4_SSL(imap_url)
# con.login(user, password)
# # print(con.list())
# con.select('AutomationFolder/Houston_IM')

# print(con.select('AutomationFolder/Houston_IM'))
# # selecting the first email
# email_id_raw = str(con.select('AutomationFolder/Houston_IM'))
# email_id = email_id_raw[10:-3]
# email_id_bytes = bytes(email_id, encoding='utf8')

# # print(email_id_bytes)

# result, data = con.fetch(email_id_bytes, '(RFC822)')
# raw = email.message_from_bytes(data[0][1])
# get_attachments(raw)
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=ArcSQL;DATABASE=RiskSandbox;TRUSTED_CONNECTION=yes')
cursor = conn.cursor()

current_date = datetime.today() - timedelta(days=1)
asof_year = int(str('20' + asofyear))
asofdate = datetime(asof_year, asof_month, asof_day)

i=0
with open(file_location,'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        if i >= 1:
            us_im_margin_zero_striped = row[7].strip('-').strip('0')
            us_im_margin = str('-'+ us_im_margin_zero_striped)
            grp_code = row[1]
            data = [asofdate, grp_code, us_im_margin, 'ARCUS', 'USD']
            print(data)
            cursor.execute("INSERT INTO InitialMarginInterest(AsOf, ClearingAccount, Value, GroupCode, Currency) VALUES(?,?,?,?,?)", data)
            cursor.commit()
        i+=1
cursor.close()
conn.close()
print('done')