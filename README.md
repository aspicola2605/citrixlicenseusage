<h2>Citrix License Usage</h2>

<b>Description:</b>
Citrix licensing real time and historical usage tracking and dashboard.

For testing purposes I have included a fake license data CSV file. The name of the file is licdatafake.csv and is included. It has 6+ months of data but not anything up to the current date. The script will output any current data to a file name of your chosing

***By default the script references the fake license CSV data file***

Instructions:
1. Copy script and JSON file to a system with access to the license server.
2. Edit the JSON file using the JSONtemplate file for a reference. If you'd like to run the script with the fake data leave the <i>outputfilename</i> section of the JSON file. When you are ready to get real data this can be updated to a new name. The script will output data to the fake license CSV file as well.
3. Run the script manually or using a scheduled task. The script has a built in loop with a cycle sleep time configured in the JSON file. The loop can be turned off if you'd like to have control using a Scheduled Task.
