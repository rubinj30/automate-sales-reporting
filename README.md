# automate-sales-reporting

## Backstory 
This was my first Python project outside of using it to analyze data. 

I knew a lot of people at the company were doing a lot of manual work with pulling reports on a daily basis and started looking into how Python could help cutdown on this. I found a library that let me pull Salesforce data in Python, and then I did some basic analytics and had a beautiful table just like the Excel table I had built for years earlier. BUT, i couldn't figure out how to send out these new tables, especially if it was going to be done automatically (My new web-dev skills would have been great for this). I had to find another way, so I was able to automatically input the data into the existing Excel reporting files. Then I had to figure out how to make API calls to our internal calling platform for call data, and put it in the same Excel file. I mapped the internal Shared drive to my desktop, so I could easily save to it in the script, then used a library to send out notifications to the specific supervisor(s) that handled each of the team's reports. 

While all this seems pretty easy looking back, it was difficult at the time. But, it was an incredible feeling to solve each and every problem that arose, which I did with almost no direction other than some advice on using the dev tools to find the URL for the call data. 

In the end, this and similar scripts were being run automatically each morning by Windows Task Manager. It replaced the daily (and in some cases bi-daily) reporting responsibilities for 4 different sales managers/supervisors each pulling anywhere from 3 to 8 different reports from 2 different sources. The total number of sales reps reported on was over 100. 

## Refactor
Once I learned about functional programming, I came back and refactored using the new skills I'd learned. Before I left the company, I setup the scripts to run on a single computer and showed each supervisor how to add and remove sales reps, and run manually if needed. As far as I know, these scripts are still being used today. 

