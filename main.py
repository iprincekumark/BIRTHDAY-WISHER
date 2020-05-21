import pandas as pd  # pip install pandas
import datetime
import smtplib  # pip install secure-smtplib
import os

# Enter your Authentication Details
os.chdir(r" ")  # Enter your Excel file directory

GMAIL_ID = ''  # Enter your Gmail Id
GMAIL_PSWD = ''  # Enter your Gmail Password


def sendEmail(to, sub, msg):
    print(
        f"Email is sent to {to} with subject: {sub} and message {msg} is sent")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)

    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


if __name__ == "_main_":
    df = pd.read_excel("data.xlsx")

    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")

    writeInd = []
    for index, item in df.iterrows():

        bday = item['Birthday'].strftime("%d-%m")

        if(today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)

    df.to_excel('data.xlsx', index=False)
