#Importing all modules
import pandas as pd
import smtplib
import datetime
import time
from email.message import EmailMessage
from getpass import getpass
sender_email = input("Enter Sender Email ID:")
sender_password = getpass("Enter Sender mail's password:")
def sendEmail(to,sub,msg):
    try:
        #connection establishment
        s = smtplib.SMTP("smtp.gmail.com",587)
        #session start
        s.starttls()
        #login with credentials
        s.login(sender_email,sender_password)
        #sending email
        s.sendmail(sender_email,to,f"Subject:{sub}\n\n{msg}")
        #session closed
        s.quit()
        print("Email sent to" + str(to) + "with subject" + str(sub) + "and message:" + str(msg))
    except Exception as ex:
        print("Ooops!Something went wrong....",ex)

if __name__=="__main__":
    #reading the excel file containing all details
    df = pd.read_excel("employees.xlsx")
    #today date in DD-MM format
    today = datetime.datetime.now().strftime("%d-%m")
    #Current year in YY format
    year = datetime.datetime.now().strftime("%Y")
    #Writing into index
    ToList =[]
    for index,item in df.iterrows():
        msg = "Happy Birthday to dear" + str(item['Employee_Name'])
        #stripping birthday in excel in DD-MM
        birthday = item['Birthday'].strftime("%d-%m")
        if (today == birthday):
          #function call
            sendEmail(item['Email_ID'],"Happy Birthday",msg)
            ToList.append(index)
