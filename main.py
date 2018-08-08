'''
Authors: Zineb & Monsef ALAHEM
sumerize: the progam send a message to all emails in excel file
'''

#excel tools
import xlrd
from xlrd import open_workbook

#auto email tools
from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL

#loading excel file
wb = open_workbook ("my_excel_file.xlsx")
sheet = wb.sheet_by_index(0)

#check numbers of rows
print(sheet.nrows)

# qq mail sending server
host_server = 'email_host_server' #exmeple 'smtp.exmail.qq.com'
sender_mail = 'my_email'
sender_passcode = 'my_password'


def send_mail(receiver='', mail_title='', mail_content=''):
    # ssl login
    smtp = SMTP_SSL(host_server, 465)
    print("ssl session success !")
    # set_debuglevel() for debug, 1 enable debug, 0 for disable
    # smtp.set_debuglevel(1)
    smtp.ehlo(host_server)
    print("ehlo success !")
    smtp.login(sender_mail, sender_passcode)
    print("login success !")

    # construct message
    msg = MIMEText(mail_content, "plain", 'utf-8')
    msg["Subject"] = Header(mail_title, 'utf-8')
    msg["From"] = sender_mail
    msg["To"] = receiver
    smtp.sendmail(sender_mail, receiver, msg.as_string())
    smtp.quit()

# loop over the excel file's rows
for rownum in range(sheet.nrows -1):

	#retriev data from the first collum, in this case the name
	name = int(sheet.cell(rownum+1,1).value)

	#retriev the data from the second collum, in this case the email
	receiver_email = int(sheet.cell(rownum+1,2).value)

	# receiver mail
	receiver = receiver_email
	# mail contents
	mail_content = "Hello,\n\nHYour name is" + str(name)
	# mail title
	mail_title = 'say hello'

	# print(sheet.cell(rownum+1,12).value)
	# print(int(sheet.cell(rownum+1,2).value))
	#print(mail_content)
	send_mail(receiver=receiver,mail_title=mail_title,mail_content=mail_content)