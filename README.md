<h2>Citrix License Usage</h2>

<b>Description:</b>
Citrix licensing real time and historical usage tracking and dashboard.

![](https://raw.githubusercontent.com/aspicola2605/citrixlicenseusage/master/img/realtime.png)
![](https://raw.githubusercontent.com/aspicola2605/citrixlicenseusage/master/img/last6.png)

For testing purposes I have included a fake license data CSV file. The name of the file is licdatafake.csv and is included. It has 6+ months of data but not anything up to the current date. The script will output any current data to a file name of your chosing

***Graphs will NOT be created in full for time ranges until enough data is gathered***

Instructions:
1. Copy script and JSON file to a system with access to the license server.
2. Install Microsoft Chart Controls for Microsoft .NET Framework 3.5 on the hosting server. This is only necessary on the script hosting server to create the charts which then output to PNG files for use on the website.
https://www.microsoft.com/en-us/download/details.aspx?id=14422
3. Edit the JSON file using the JSONtemplate file for a reference. If you'd like to run the script with the fake data leave the <i>outputfilename</i> section of the JSON file. When you are ready to get real data this can be updated to a new name. The script will output data to the fake license CSV file as well.
4. Run the script manually or using a scheduled task. The script has a built in loop with a cycle sleep time configured in the JSON file. The loop can be turned off if you'd like to have control using a Scheduled Task.

***All data is fake in included datafile and screenshot. All files have been generalized***
