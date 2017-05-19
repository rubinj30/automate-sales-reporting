## Pull Closed Deals from Salesforce
import pandas as pd
import numpy as np
from salesforce_reporting import Connection, ReportParser

#Connects to Salesforce
sf = Connection(username='myWork@email.com', password='myPassword', security_token='myToken')

# Pull individual Outbound Closed Deals (just rep name and $ amount for each deal)
sfReportDeals = sf.get_report('00O14000008ueQW', details=True)
dealsParser = ReportParser(sfReportDeals)
dealsList = dealsParser.records()

#turns List into DataFrame
dealsDF = pd.DataFrame(dealsList)
dealsDF.columns = ['Account Owner', 'Daily']

#Changes MRC column from string to float
dealsDF['Daily'] = pd.to_numeric(dealsDF['Daily'], errors='coerce').fillna(0)

#Totals by Sales Rep
dealsByRepDF = dealsDF.groupby('Account Owner')

dollarsAndDeals = dealsByRepDF.agg([np.sum, np.count_nonzero]).xs('Daily', axis=1, drop_level=True).reset_index()

# Pull individual new Opportunities

oppsReport = sf.get_report('00O14000008uelo', details=True)
oppsParser = ReportParser(oppsReport)
oppsList = oppsParser.records()
oppsDF = pd.DataFrame(oppsList)

# Column starts as 0 so renaming it Opportunity Owner
oppsDF.columns = ['QLs']

# Count Opps
oppsDFCount = oppsDF['QLs'].value_counts()
oppsDFCount2 = pd.DataFrame(oppsDFCount)
oppsDFCount2.reset_index(inplace=True)
oppsDFCount2.rename(columns={'index':'Account Owner','QLs':'# of QLs'},inplace=True)


# combine $'s and # of 
dollarsAndDeals.columns = ['Account Owner',"Total $'s", "# of Deals"]
dollarsDealsQLs = dollarsAndDeals.merge(oppsDFCount2, on='Account Owner')

# remove column
dollarsDealsQLs.columns = ['Account Owner',"Total $'s",'# of Deals','# of QLs']

    # today for HDAP
todayMonth = datetime.datetime.now().strftime('%m')
todayDay = datetime.datetime.now().strftime('%d')

    #yesterday for HDAP
yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
yesterdayMonth = yesterday.strftime('%m')
yesterdayDay = yesterday.strftime('%d')
    
    # JSON call data from HDAP
url = 'https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-' + yesterdayMonth + '-' + yesterdayDay + 'T00:00:00Z&endDateInGMT=2017-'+todayMonth+'-'+todayDay+'T00:00:00Z&lineChartType=Total%20Calls&accountId=25016'
stuff = requests.get(url, auth=(hdapUsername, password)).content
cleaner_stuff = str(stuff, 'utf-8')
cleanest_stuff = json.loads(cleaner_stuff)

# Puts call data for specified reps into a list of lists
callDataList = []
def pullHDAPData(person):
    for i in range(1040):
        salesRep = cleanest_stuff[i]['category']['categoryName']
        calls = int(cleanest_stuff[i]['metrics'][4]['value'])
        talkTime = cleanest_stuff[i]['metrics'][1]['value']
      
        if salesRep == person:
            individualCallData = [salesRep, calls, talkTime] #datetime.time(*map(int, talkTime.split(':')))]
            callDataList.append(individualCallData)
            
names = ['Jay Van Horn','David New','Derek Kirchner','Oliver Haney','Matthew Elliott','Kirk Sweeney', 'Dylan Stephens','Marquis Jacox','Jacqueline Momeni']

for name in names:
	pullHDAPData(name)

callDataDF = pd.DataFrame(callDataList)
callDataDF.columns = ['Account Owner', 'Follow-ups','Total Talk Time']


allDataDF = dollarsDealsQLs.merge(callDataDF, on='Account Owner')
#htmlDF = allDataDF.to_html()
