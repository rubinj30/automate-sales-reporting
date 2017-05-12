import requests
import json
import openpyxl as xl
from salesforce_reporting import Connection, ReportParser
import datetime
import smtplib


# Asks for user info to verify 
computer = input("What computer are you on?")
username = input("What is your HDAP admin username?")
sfEmail = input("What is your salesforce e-mail?")
password = input("What is your salesforce password?")
gmailPassword = input("What is your gmail password?")

# SF connection
sf = Connection(username=sfEmail, password=password, security_token='myToken')


# if today is not Monday (0), then run salesforce reports for yesterday. 
# else (or 1st workday), will pull different salesforce reports that

if datetime.datetime.today().weekday() != 0:
    
    # today for HDAP
    todayMonth = datetime.datetime.now().strftime('%m')
    todayDay = datetime.datetime.now().strftime('%d')

    #yesterday for HDAP
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    yesterdayMonth = yesterday.strftime('%m')
    yesterdayDay = yesterday.strftime('%d')
    
    # JSON call data from HDAP
    url = "https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-' + yesterdayMonth + '-' + yesterdayDay + 'T00:00:00Z&endDateInGMT=2017-'+todayMonth+'-'+todayDay+'T00:00:00Z&lineChartType=Total%20Calls&accountId='acctNumber'""
    stuff = requests.get(url, auth=(username, password)).content
    cleaner_stuff = str(stuff, 'utf-8')
    cleanest_stuff = json.loads(cleaner_stuff)


# Searches JSON for the names in for loop and spits out 'Outbound calls' (as integers) and 'Total Talk Time'
    for i in range(1040):

        salesRep = cleanest_stuff[i]['category']['categoryName']
        calls = int(cleanest_stuff[i]['metrics'][4]['value'])
        talkTime = cleanest_stuff[i]['metrics'][1]['value']

        if salesRep == 'Jay Van Horn':
            jayList = [salesRep, calls, datetime.time(*map(int, talkTime.split(':')))]
        if salesRep == 'David New':
            davidList = [salesRep, calls, datetime.time(*map(int, talkTime.split(':')))]
        if salesRep == 'Derek Kirchner':
            derekList = [salesRep, calls, datetime.time(*map(int, talkTime.split(':')))]
        if salesRep == 'Oliver Haney':
            oliverList = [salesRep, calls, datetime.time(*map(int, talkTime.split(':')))]


    # Need to append this list of lists to Excel sheet 'HDAP'               
    totalList = [davidList, derekList, jayList, oliverList]

# New Outbound Opps from Salesforce
    yesterdayOpps = '00O14000008ywre'
    oppsSf = sf.get_report(yesterdayOpps, details=True)
    oppsParser = ReportParser(oppsSf)
    opps = oppsParser.records()

    
# Closed Organic Opps from Salesforce
    yesterdayOrganic = '00O14000008yuxL'
    organicSf = sf.get_report(yesterdayOrganic, details=True)
    organicParser = ReportParser(organicSf)
    organic = organicParser.records()
    for value in organic:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])

# Closed TLC Opps from Salesforce
    yesterdayTlc = '00O14000008yuzW'
    tlcSf = sf.get_report(yesterdayTlc, details=True)
    tlcParser = ReportParser(tlcSf)
    tlc = tlcParser.records()
    for value in tlc:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])

# Closed Convergys Opps from Salesforce
    yesterdayConvergys = '00O14000008yuxV'
    convergysSf = sf.get_report(yesterdayConvergys, details=True)
    convergysParser = ReportParser(convergysSf)
    convergys = convergysParser.records()
    for value in convergys:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])
    

    # Open Excel template and insert the lists of data for yesterday
    xlOpen = xl.load_workbook('path/OB-ISR.xlsx')

    # Get HDAP sheet and insert Reps, Calls, and Talk Time
    hdapSheet = xlOpen.get_sheet_by_name('HDAP')
    for data in totalList:
        hdapSheet.append(data)


    #Get opps sheet and insert new Opportunities
    oppsSheet = xlOpen.get_sheet_by_name('opps')
    for data in opps:
        oppsSheet.append(data)

    #Get organic sheet and insert organic deals
    organicSheet = xlOpen.get_sheet_by_name('organic')
    for data in organic:
        organicSheet.append(data)

#Get tlc sheet and insert tlc deals
    tlcSheet = xlOpen.get_sheet_by_name('tlc')
    for data in tlc:
        tlcSheet.append(data)

#Get convergys sheet and insert convergys deals
    convergysSheet = xlOpen.get_sheet_by_name('convergys')
    for data in convergys:
        convergysSheet.append(data)


    # Pull date from existing variable for label of Excel file
    yesterdayLabel = yesterday.strftime('%B-%d')

    # Save Excel file in shared Google Drive folder
    xlOpen.save('/path/' + yesterdayLabel + ' OB-ISR.xlsx')
    
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()


    smtpObj.login('jonathan.rubin@vonage.com', gmailPassword)
#can use list of strings for multiple email addresses. save as object and replace singular email

    hyperlinkDrive = 'https://drive.google.com/drive/folders/link'
    recipients = ['recipient@email.com','recipient2@email.com','recipient3@email.com']
    smtpObj.sendmail('from@email.com',recipients,
	'Subject: Daily Report\nHello,\n\nYou can find the completed Outbound closer report in the Google Drive folder:\n\n' + hyperlinkDrive + '\n\nTry opening with Excel or Google Sheets or just downloading and then opening with Excel.\n\nThanks,\nJonathan')

#disconnects from server
    smtpObj.quit()


else:

    # today for HDAP
    saturday = datetime.datetime.now() - datetime.timedelta(days=2)
    saturdayMonth = datetime.datetime.now().strftime('%m')
    saturdayDay = datetime.datetime.now().strftime('%d')

    #yesterday for HDAP
    friday = datetime.datetime.now() - datetime.timedelta(days=3)
    fridayMonth = yesterday.strftime('%m')
    fridayDay = yesterday.strftime('%d')
    
    # JSON call data from HDAP
    url = 'https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-' + fridayMonth + '-' + fridayDay + 'T00:00:00Z&endDateInGMT=2017-'+saturdayMonth+'-'+saturdayDay+'T00:00:00Z&lineChartType=Total%20Calls&accountId=25016'
    stuff = requests.get(url, auth=(username, password)).content
    cleaner_stuff = str(stuff, 'utf-8')
    cleanest_stuff = json.loads(cleaner_stuff)


# Searches JSON for the names in for loop and spits out 'Outbound calls' (as integers) and 'Total Talk Time'
    for i in range(1040):

        salesRep = cleanest_stuff[i]['category']['categoryName']
        calls = int(cleanest_stuff[i]['metrics'][4]['value'])
        talkTime = cleanest_stuff[i]['metrics'][1]['value']

        if salesRep == 'Jay Van Horn':
            jayList = [salesRep, calls, talkTime]
        if salesRep == 'David New':
            davidList = [salesRep, calls, talkTime]
        if salesRep == 'Derek Kirchner':
            derekList = [salesRep, calls, talkTime]
        if salesRep == 'Oliver Haney':
            oliverList = [salesRep, calls, talkTime]


    # Need to append this list of lists to Excel sheet 'HDAP'               
    totalList = [davidList, derekList, jayList, oliverList]

# New OB Opps from Salesforce    
    manualOpps = '00O14000008ywrj'
    oppsSf = sf.get_report(manualOpps, details=True)
    oppsParser = ReportParser(oppsSf)
    opps = oppsParser.records()

    
# Closed Organic Opps from Salesforce
    manualOrganic = '00O14000008yuzM'
    organicSf = sf.get_report(manualOrganic, details=True)
    organicParser = ReportParser(organicSf)
    organic = organicParser.records()
    for value in organic:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])

# Closed TLC Opps from Salesforce
    manualTlc = '00O14000008yuzb'
    tlcSf = sf.get_report(manualTlc, details=True)
    tlcParser = ReportParser(tlcSf)
    tlc = tlcParser.records()
    for value in tlc:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])

# Closed Convergys Opps from Salesforce
    manualConvergys = '00O14000008yuzg'
    convergysSf = sf.get_report(manualConvergys, details=True)
    convergysParser = ReportParser(convergysSf)
    convergys = convergysParser.records()
    for value in convergys:
        value[4] = value[4].replace("$", "")
        value[4] = float(value[4])

    # Open Excel template and insert the lists of data for yesterday      
    xlOpen = xl.load_workbook('/path/OB-ISR.xlsx')

    # Get Excel 'HDAP' sheet and insert Reps, Calls, and Talk Time
    hdapSheet = xlOpen.get_sheet_by_name('HDAP')
    for data in totalList:
        hdapSheet.append(data)

    #Get Excel 'opps' sheet and insert new Opportunities
    oppsSheet = xlOpen.get_sheet_by_name('opps')
    for data in opps:
        oppsSheet.append(data)

    #Get Excell 'organic' sheet and insert organic deals
    organicSheet = xlOpen.get_sheet_by_name('organic')
    for data in organic:
        organicSheet.append(data)

    #Get Excel 'tlc' sheet and insert tlc deals
    tlcSheet = xlOpen.get_sheet_by_name('tlc')
    for data in tlc:
        tlcSheet.append(data)

    #Get Excel 'convergys' sheet and insert convergys deals
    convergysSheet = xlOpen.get_sheet_by_name('convergys')
    for data in convergys:
        convergysSheet.append(data)


    # Pull date from existing variable for label of Excel file    
    fridayLabel = friday.strftime('%B-%d')

    # Save file to shared Google Drive folder
    xlOpen.save('/path/' + fridayLabel + ' OB-ISR.xlsx')
    
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()

#should call input to require password to be entered

    smtpObj.login('jonathan.rubin@vonage.com', gmailPassword)
#can use list of strings for multiple email addresses. save as object and replace singular email

    hyperlinkDrive = 'https://drive.google.com/drive/folders/link'
    recipients = ['recipient@email.com','recipient2@email.com','recipient3@email.com']
    smtpObj.sendmail('from@email.com',recipients,
	'Subject: Daily Report\nHello,\n\n You can find the updated report in the Google Drive folder:\n\n' + hyperlinkDrive + '\n\nTry opening with Excel or Google Sheets or just downloading and then opening with Excel.\n\nThanks,\nJonathan')

#disconnects from server
    smtpObj.quit()

