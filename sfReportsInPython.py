# Pull Closed Deals from Salesforce
import pandas as pd
import numpy as np
from salesforce_reporting import Connection, ReportParser

# Connects to Salesforce
sf_connec = Connection(username='myWork@email.com', password='myPassword', security_token='myToken')

# Pull individual Outbound Closed Deals (just rep name and $ amount for each deal)
closed_deals_report = sf_connec.get_report('00O14000008ueQW', details=True)
closed_deals_parsed = ReportParser(closed_deals_report)
closed_deals_list = closed_deals_parsed.records()

# turns List into DataFrame
closed_deals_df = pd.DataFrame(closed_deals_list)
closed_deals_df.columns = ['Account Owner', 'Daily']

# Changes MRC column from string to float
closed_deals_df['Daily'] = pd.to_numeric(closed_deals_df['Daily'], errors='coerce').fillna(0)

# Totals by Sales Rep
deals_by_rep_df = closed_deals_df.groupby('Account Owner')

dollars_and_deals_df = deals_by_rep_df.agg([np.sum, np.count_nonzero]).xs('Daily', axis=1, drop_level=True).reset_index()

# Pull individual new Opportunities

opps_report = sf.get_report('00O14000008uelo', details=True)
opps_parsed = ReportParser(opps_report)
opps_list = opps_parsed.records()
opps_df = pd.DataFrame(opps_list)

# Column starts as 0 so renaming it Opportunity Owner
opps_df.columns = ['QLs']

# Count Opps
opps_count_df = opps_df['QLs'].value_counts()
opps_count_df = pd.DataFrame(opps_count_df)
opps_count_df.reset_index(inplace=True)
opps_count_df.rename(columns={'index':'Account Owner','QLs':'# of QLs'},inplace=True)

# combine $'s and # of 
dollars_and_deals_df.columns = ['Account Owner',"Total $'s", "# of Deals"]
dollars_deals_quals_df = dollars_and_deals_df.merge(opps_count_df, on='Account Owner')

# remove column
dollars_deals_quals_df.columns = ['Account Owner',"Total $'s",'# of Deals','# of QLs']

# today for HDAP
today_month = datetime.datetime.now().strftime('%m')
today_day = datetime.datetime.now().strftime('%d')

# yesterday for HDAP
yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
yesterday_month = yesterday.strftime('%m')
yesterday_day = yesterday.strftime('%d')
    
# JSON call data from HDAP
url = 'https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-' + yesterday_month + '-' + yesterday_day + 'T00:00:00Z&endDateInGMT=2017-'+today_month+'-'+today_day+'T00:00:00Z&lineChartType=Total%20Calls&accountId=25016'
stuff = requests.get(url, auth=(hdap_username, password)).content
cleaner_stuff = str(stuff, 'utf-8')
cleanest_stuff = json.loads(cleaner_stuff)

# Puts call data for specified reps into a list of lists
call_data_list = []
def pull_call_data(person):
    for i in range(1040):
        sales_rep = cleanest_stuff[i]['category']['categoryName']
        calls = int(cleanest_stuff[i]['metrics'][4]['value'])
        talk_time = cleanest_stuff[i]['metrics'][1]['value']
      
        if sales_rep == person:
            individual_call_data = [sales_rep, calls, talk_time] #datetime.time(*map(int, talk_time.split(':')))]
            call_data_list.append(individual_call_data)
            
names = ['Jay Van Horn','David New','Derek Kirchner','Oliver Haney','Matthew Elliott','Kirk Sweeney', 'Dylan Stephens','Marquis Jacox','Jacqueline Momeni']

for name in names:
	pull_call_data(name)

call_data_df = pd.DataFrame(call_data_list)
call_data_df.columns = ['Account Owner', 'Follow-ups','Total Talk Time']

all_data_df = dollars_deals_quals_df.merge(call_data_df, on='Account Owner')
