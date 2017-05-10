import requests
import json
url = "https://my.vonagebusiness.com/appserver/rest/usersummarymetricsresource?chartType=summary&startDateInGMT=2017-5-8T00:00:00Z&endDateInGMT=2017-5-9T00:00:00Z&lineChartType=Total%20Calls&accountId='AccountNumber'"
username = 'myUsername'
password = 'myPassword'
stuff = requests.get(url, auth=(username, password)).content
cleaner_stuff = str(stuff, 'utf-8')
cleanest_stuff = json.loads(cleaner_stuff)

# For reading JSON in case i want to include other data from HDAP
#print(json.dumps(cleanest_stuff, indent=4, sort_keys=True))

# indexes in JSON for all OB and SPG sales reps
#OB = [597,244,222,60,184,982,694,109]

# names of OB only reps
#OBnames = ['Jay Van Horn','Derek Kirchner','David New','Oliver Haney']

for i in range(1040):
    salesRep = cleanest_stuff[i]['category']['categoryName']
    calls = cleanest_stuff[i]['metrics'][4]['value']
    talkTime = cleanest_stuff[i]['metrics'][1]['value']
    if salesRep == 'Jay LastName':
        jayList = [salesRep, calls, talkTime]
    if salesRep == 'David LastName':
        davidList = [salesRep, calls, talkTime]
    if salesRep == 'Derek LastName':
        derekList = [salesRep, calls, talkTime]
    if salesRep == 'Oliver LastName':
        oliverList = [salesRep, calls, talkTime]

# Need to append this list of lists to Excel sheet 'HDAP'               
totalList = [davidList, derekList, jayList, oliverList]
print(totalList)
