import yagmail
import pandas
from news import NewsFeed
import datetime
import time


def send_email():

    today = datetime.datetime.now().strftime('%Y-%m-%d')
    yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    news_feed = NewsFeed(interest=row['interest'],
                         from_date=yesterday,
                         to_date=today)
    email = yagmail.SMTP(user='your email', password='your password')   # you might need to generate a password for third-party apps from email settings(it worked for me to use E-mail and Computer with Windows parameters)
    email.send(to=row['email'],
               subject=f"Your {row['interest']} news for today",
               contents=f"Hi {row['name']}\n See what is on about {row['interest']} today.\n {news_feed.get()}\n Bastaman")


while True:

    if datetime.datetime.now().hour == 16 and datetime.datetime.now().minute == 21:   # This app runs 24/7, change the hour and minute variables so you will receive the emails

        df = pandas.read_excel('people.xlsx')   # I used dropmail.me to generate the emails, the one from the 'people.xlsx' don't work anymore

        for index, row in df.iterrows():
            send_email()
    time.sleep(60)