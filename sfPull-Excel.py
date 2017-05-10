
import openpyxl as xl
from salesforce_reporting import Connection, ReportParser
import datetime
import smtplib

# change this based on the computer it is being ran from ('compName1','compName2')
computer = input("What computer are you on?")
#dayOfWeek = input("Was yesterday a workday? Enter 'no' if not.")
#pword = input("What is your password")

sf = Connection(username='my@email.com', password='myPassword', security_token='myToken')

# New Opportunities from Salesforce (2 reports since Salesforce doesn't allow for '3 days prior' for Monday)

# if today is not Monday (0), then run salesforce reports for yesterday. 
# else, will pull sf reports with manual

if datetime.datetime.today().weekday() != 0:
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
    xlOpen = xl.load_workbook('/Users/' + computer + '/Dropbox/python/workAutomation/OB-ISR.xlsx')

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
    
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    yesterday = yesterday.strftime('%B-%d')
    xlOpen.save('/Users/' + computer + '/Google Drive/OB Daily Reports/' + yesterday + ' OB-ISR.xlsx')
    
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()

#should call input to require password to be entered

    smtpObj.login('my@email.com', 'myPassword')
#can use list of strings for multiple email addresses. save as object and replace singular email

    hyperlinkDrive = 'https://drive.google.com/drive/folders/0ByBUMjHVt7glV042SU96ZHhqTEU'
    recipients = ['email1@address.com','email2address.com']
    smtpObj.sendmail('jonathan.rubin@vonage.com',recipients,
	'Subject: Daily Report\nHello,\n\nYou can find the updated report in the Google Drive folder:\n\n' + hyperlinkDrive + '\n\nThanks,\nJonathan')

#disconnects from server
    smtpObj.quit()


else:
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
    xlOpen = xl.load_workbook('/Users/' + computer + '/Dropbox/python/workAutomation/OB-ISR.xlsx')

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

    friday = datetime.datetime.now() - datetime.timedelta(days=1)
    friday = friday.strftime('%B-%d')
    xlOpen.save('/Users/' + computer + '/Google Drive/OB Daily Reports/'+ friday + ' OB-ISR.xlsx')
    
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()

#should call input to require password to be entered

    smtpObj.login('my@email.com', 'myPassword')
#can use list of strings for multiple email addresses. save as object and replace singular email

    hyperlinkDrive = 'https://drive.google.com/drive/folders/0ByBUMjHVt7glV042SU96ZHhqTEU'
    recipients = ['email1@address.com','email2address.com']
    smtpObj.sendmail('my@email',recipients,
	'Subject: Daily Report\nHello,\n\n You can find the updated report in the Google Drive folder:\n\n' + hyperlinkDrive + '\n\n From that link, you can open in with Google Sheets or Excel. Or you can just download and then open with Excel.\n\nThanks,\nJonathan')

#disconnects from server
    smtpObj.quit()
