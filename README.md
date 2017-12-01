# automate-sales-reporting

## Purpose
This runs each monrning automatically and e-mails the specified Supervisors with their reports to send out to their teams. The formulas are done in Excel, so that Supervisors can add/remove items if the Salesforce data is incorrect prior to sending. 

## How it works
It uses a salesforce API to pull the previous days qualified leads taken and deals closed, as well a Vonage Business calling platform API for call activity.

The salesforce data is placed into the corresponding Excel sheets, which Excel formulas pull from. The call data comes back in JSON and is looped through to find call data for specified reps, and is then placed into corresponding Excel sheets. 

Then an e-mail is automatically sent to the Supervisor with a link to the Google Drive where they can find the report for that day. 

