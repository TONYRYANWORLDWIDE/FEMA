#1234
import requests
import urllib
from urllib import *
import pandas as pd
import os
import smtplib,ssl
from smtplib import SMTP_SSL as SMTP 
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

path = 'C:/Users/Tonyryanworldwide/Untitled Folder/fema'
os.chdir(path)
cwd = os.getcwd()

req = requests.get("https://www.fema.gov/MEDIA-LIBRARY/ASSETS/documents/132213")
req_text = req.text
start = req_text.find('https://www.fema.gov/media-library-data')
end = req_text.find('xlsx', start)
file = req_text[start:end + 4]
FEMA = pd.read_excel(file, sheet_name = 2, header = 2, convert_float = True)
FEMA = FEMA.dropna()
FEMA = FEMA[FEMA['Zip'] != 0.0]
FEMA['Zip'] = FEMA['Zip'].astype(int)
FEMA['Zip'] = FEMA['Zip'].apply(str)
FEMA['Zip'] = FEMA['Zip'].apply(lambda x: x.zfill(5))

excel_file = 'FEMA.xlsx'
writer = pd.ExcelWriter(excel_file)
FEMA.to_excel(writer, sheet_name = 'FEMA')
HouseHolds = pd.read_csv('Census_Pop_Housholds_By_ZCSTA.csv')
HouseHolds.to_excel(writer, sheet_name = 'Census Data')
FEMA_GROUP = FEMA.groupby(['Disaster Number' , 'Zip' , 'State'] , as_index = False).agg({'Home Damage' :'sum'})
HouseHolds['name'][0][-5:]
HouseHolds['ZIP'] = HouseHolds['name'].apply(lambda x: [x][0][-5:])
Final = HouseHolds.merge(FEMA_GROUP, how = 'inner', left_on = 'ZIP', right_on = 'Zip')[['Zip','Total Pop', 'arealand',  'Housing Unit Count' ,  'Disaster Number' , 'State', 'Home Damage']]
Final['Freq_Ratio'] = Final['Home Damage'] / Final['Housing Unit Count']
Final_By_Disaster = Final[['Disaster Number', 'Housing Unit Count' , 'Home Damage', 'Total Pop']].groupby('Disaster Number', as_index = False).sum()
Final_By_Disaster['Dis_Freq'] = Final_By_Disaster['Home Damage'] / Final_By_Disaster['Housing Unit Count']
Final_By_Disaster['Pop_Freq'] = Final_By_Disaster['Home Damage'] / Final_By_Disaster['Total Pop']
Final.to_excel(writer, sheet_name = 'Frequency_Ratio')
Final_By_Disaster.to_excel(writer, sheet_name = 'Disaster Ratios')
workbook  = writer.book
worksheet = writer.sheets['FEMA']
worksheet.set_column('B:Z', 20)
worksheet = writer.sheets['Frequency_Ratio']
worksheet.set_column('B:Z', 20)
worksheet = writer.sheets['Disaster Ratios']
worksheet.set_column('B:Z', 20)
writer.save()
writer.close()

file = 'C:/Users/Tonyryanworldwide/Untitled Folder/fema/FEMA.xlsx'
username=''
password=''
send_from = ''
send_to = ''
Cc = ''
msg = MIMEMultipart()
msg['From'] = send_from
msg['To'] = send_to
msg['Cc'] = Cc
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = 'FEMA FREQ'
server = smtplib.SMTP('smtp.gmail.com')
port = '587'
fp = open(file, 'rb')
part = MIMEBase('application','vnd.ms-excel')
part.set_payload(fp.read())
fp.close()
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename='FEMA FREQ')
msg.attach(part)
smtp = smtplib.SMTP('smtp.gmail.com')
smtp.ehlo()
smtp.starttls()
smtp.login(username,password)
smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
smtp.quit()

