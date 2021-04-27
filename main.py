import pandas as pd
import datetime
import smtplib

GMAIL_ID = ''
GMAIL_PSWD = ''


def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


if __name__ == "__main__":

    df = pd.read_excel("data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    writeIndex = []
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        if (today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'], item['Subject'], item['Message'])
            writeIndex.append(index)

    for i in writeIndex:
        year = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(year) + ', ' + str(yearNow)

    df.to_excel('data.xlsx', index=False)
