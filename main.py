import datetime
import pandas as pd
from datetime import date, datetime, timedelta
from pandas import DataFrame
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from win32com.client import Dispatch

FILENAME = r"C:\Users\royno\OneDrive\Desktop\example.xlsx"
CHECK_KEYS = ['תאריך תפוגה', 'תפוגה טופס']
DATE_FORMAT = "%d-%m-%Y"


def read_excel():
    custom_date_parser = lambda x: datetime.strptime(x, DATE_FORMAT)
    xls = pd.ExcelFile(
        FILENAME,
        engine='openpyxl',
    )
    df = pd.read_excel(xls, None, parse_dates=CHECK_KEYS, date_parser=custom_date_parser)
    return df


def filter_soon_to_expired(df, sheetname):
    if 'קשר' in sheetname:
        return
    today_str = date.today().strftime(DATE_FORMAT)
    today_date = datetime.strptime(today_str, DATE_FORMAT)
    alert_on_date = today_date + timedelta(days=30)
    filtered = df.loc[(df[CHECK_KEYS[0]] < alert_on_date) | (df[CHECK_KEYS[1]] < alert_on_date)]
    return filtered


def send_gmail(contacts, alerts, department):
    print(f'Sending e-mail to {contacts}')
    mail_content = f"""Hello,
    I wish to inform you there are less than 30 days until expired for the below medicine
    
    {alerts}
    """
    sender_address = 'roynovich1451@gmail.com'
    sender_pass = '08R09n1991'
    receiver_address = contacts
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Notice - Expiration date is close'
    message.attach(MIMEText(mail_content, 'plain'))
    session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
    session.starttls()  # enable security
    session.login(sender_address, sender_pass)  # login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent')


def send_outlook(contacts, alerts, department):
    valid_contacts = ';'.join(contacts.split(','))  # recipients mails must be separate by ';'
    mail_content_html = f"""<h2>Hello,
    I wish to inform you there are less than 30 days until expired for the below medicine</h2>

    {alerts.to_html()}
    """
    outlook = Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0x0)
    mail.To = valid_contacts
    mail.Subject = 'Notice - Expiration date is close'
    mail.HTMLBody = mail_content_html

    mail.Send()
    print('Mail Sent')


def get_contacts(contacts_sheet: DataFrame, department):
    df_contacts = contacts_sheet.loc[contacts_sheet['מחלקה'] == department]
    df_contacts_remove_empty = df_contacts.dropna(axis='columns')
    return df_contacts_remove_empty.to_numpy()[0][1]


def main():
    df = read_excel()
    for department, sheet in df.items():
        if 'אנשי קשר' in department:
            continue
        df_soon_exp = filter_soon_to_expired(sheet, department)
        if not df_soon_exp.empty:
            send_outlook(get_contacts(df['אנשי קשר'], department), df_soon_exp, department)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/


# def convert_date(date_str):
#     print(date_str)
#     splited = re.split('/|\.', date_str)
#     print(splited)
#     if len(splited[0]) != 2:
#         splited[0] = f'0{splited[0]}'
#     if len(splited[1]) != 2:
#         splited[1] = f'0{splited[1]}'
#     if len(splited[2]) != 4:
#         splited[2] = f'20{splited[2]}'
#     date_str = '/'.join(splited)
#     return date_str

# def evaluate_date(datecheck):
#
#     print(f'ROYYY {datecheck}')
#     today_str = date.today().strftime(DATE_FORMAT)
#     today_date = datetime.strptime(today_str, DATE_FORMAT)
#     date_str = convert_date(datecheck)
#     check_date = datetime.strptime(date_str, DATE_FORMAT)
#     delta = check_date - today_date
#     if delta.days <= 30:
#         return True
#     return False

# def search_expires(sheet, sheetname):
#     alert_rows = {}
#     if 'קשר' in sheetname:
#         return
#     js_string = sheet.to_json(force_ascii=False)
#     js = json.loads(js_string)
#
#     check_keys = [k for k in js.keys() if 'תפוגה' in k]
#     for key in check_keys:
#         for index in js[key].keys():
#             expired = evaluate_date(js[key][index])
#             if key not in alert_rows.keys():
#                 alert_rows[key] = [index]
#             else:
#                 alert_rows[key].append(index)
#
#     print(f'expired rows = {alert_rows}')
