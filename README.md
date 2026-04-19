# FundManager I Bond Quote Server
This is a powershell script that downloads I Bond pricing from Treasury Direct and formats it into a csv that FundManager will ingest as a <local file> price server.

To use either setup you will need to:
1) Point the <local file> Internet Retrieve Option to the generated ibondquoteserver.csv in this folder
2) Set your I-Bonds in FM to the following ticker format IB + MON + YYYY -- for example April 2023 I-bond would be IBAPR2023
3) Set your I-Bonds in FM to use the <local file> option for current and historical prices 
4) Set your I Bond starting price at $100 and make the number of shares the mutiple of $100 -- for example if you bought a $1253 I Bond you would set the share price to $100 and the # of shares to 12.53

There are 2 ways to use it / this repository:

# Download Completed CSV
This github repo automatically runs the script and updates the IbondFM-QuoteServer.csv file included every May and Novemember 4th.  You can manually download the file after the 4th OR

You can import the included IBond-FMPriceDownload.xml into Scheduled Tasks in Windows to automatically download the file on May and November 5th automatically. By default it will store the output csv in the following folder:

C:\FundManager-CustomQuoteServers

You can either create that folder OR edit the task's stating directory to match the location you'd like the file to be downloaded to.

If you'd like to view the download command that is being used you can look either in the batch file or in the scheduled task and it will show you what it is running. (Just curl to the right github path)

# Run Locally
To run the script locally you will have to set the execution policy to allow scripts. And the script will store the file in the folder the shell is runnnig in.  You can aslo schedule a task to run this if you'd like.

# That's It!
Once the following is done your I Bond prices should update by the 5th of May and November for the next 6 months. The slight delay is to account for weekends and holidays. So failures don't just not update the price file.  
