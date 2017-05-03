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

# final chart of with all Outbound reps with their
# (1) total number of new Opportunities
# (2) total number of closed Opportunities/Deals
# (3) total dollar amount of all their closed Opportunities/Deals
dollarsDealsQLs
