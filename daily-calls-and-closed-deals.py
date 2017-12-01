import requests
import json
import openpyxl as xl
from salesforce_reporting import Connection, ReportParser
import datetime
import smtplib
from credentials import login

sf = Connection(username=login['sfEmail'], password=login['password'], security_token=login['token'])

excel_report_template = xl.load_workbook('/Users/rubinj30/Google Drive/OB Sales Reports/OB-ISR-template.xlsx')

# get variables for API call and labeling Excel files
today_month = datetime.datetime.now().strftime('%m')
today_day = datetime.datetime.now().strftime('%d')
today_month_day = today_month + today_day
yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
yesterday_month = yesterday.strftime('%m')
yesterday_day = yesterday.strftime('%d')
yesterday_year = yesterday.strftime('%Y')
yesterday_month_day = yesterday_month + '-' + yesterday_day

call_data_list = []
names = ['first_name last_name', 'first_name last_name', 'first_name last_name', 'first_name last_name', 'first_name last_name']
url_for_api_call = f'https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-{yesterday_month_day}T00:00:00Z&endDateInGMT=2017-{today_month+}-{today_day}T00:00:00Z&lineChartType=Total%20Calls&accountId=25016'
json_call_data = requests.get(url_for_api_call, auth=(login['username'], login['password'])).json()


def pull_number_of_calls_and_talk_time(person):
    for i in range(len(json_call_data)):
        salesRep = json_call_data[i]['category']['categoryName']
        calls = int(json_call_data[i]['metrics'][4]['value'])
        talkTime = json_call_data[i]['metrics'][1]['value']
        if salesRep == person:
            indiv_call_data = [salesRep, calls, datetime.time(*map(int, talkTime.split(':')))]
            call_data_list.append(indiv_call_data)

for name in names:
    pull_number_of_calls_and_talk_time(name)


def append_hdap_to_excel_sheet(sheet_name):
    excel_sheet = excel_report_template.get_sheet_by_name(sheet_name)
    for data in call_data_list:
        excel_sheet.append(data)


def sf_report_to_excel_sheet(report_id, sheet_name):
    report = ReportParser(sf.get_report(report_id, details=True)).records()
    excel_sheet = excel_report_template.get_sheet_by_name(sheet_name)
    for data in report:
        excel_sheet.append(data)


def sf_report_to_sheet_and_edit_type(report_id, sheet_name):
    report = ReportParser(sf.get_report(report_id, details=True)).records()
    for value in report:
        value[4] = value[4].replace("$", "").replace("(", "").replace(")", "").replace(",", "").replace("-", "0")
        value[4] = float(value[4])
    excel_sheet = excel_report_template.get_sheet_by_name(sheet_name)
    for data in report:
        excel_sheet.append(data)


def send_email_notification_with_link(google_drive_link):
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login('jonathan.rubin@vonage.com', login['gmailPassword'])

    # can use list of strings for multiple email addresses. save as object and replace singular email
    recipients = ['recipient@email.com', 'recipient2@email.com', 'recipient3@email.com']
    smtpObj.sendmail('myemail@email.com', recipients,
                     f"Subject: Refactor Test {yesterday_month}-{yesterday_day}" +
                     f"\nHere you can find the Closer report for yesterday:\n\n{google_drive_link}" +
                     "\n\nYou can open it with Excel or Google Sheets.\n\nThanks,\nJonathan")

    # disconnects from server
    smtpObj.quit()

# specific report IDs and Excel sheet names
append_hdap_to_excel_sheet('HDAP')
sf_report_to_excel_sheet('00O14000008ywre', 'opps')
sf_report_to_sheet_and_edit_type('00O14000008yuxL', 'organic')
sf_report_to_sheet_and_edit_type('00O14000008yuzW', 'tlc')
sf_report_to_sheet_and_edit_type('00O14000008yuxV', 'convergys')
sf_report_to_sheet_and_edit_type('00O1O000009KQ8z', 'concentrix')
sf_report_to_sheet_and_edit_type('00O1O0000086fvQ', "Total MTD $'s")

# enter drive link for specific team
send_email_notification_with_link('https://drive.google.com/drive/folders/outbound')

excel_report_template.save(f'/Users/rubinj30/Google Drive/OB Daily Reports/OB ISR Daily - {today_month_day}.xlsx')
