#=========================================================================================================================
<#
    Author: Alex Spicola
    Created: March 2018
    
    Purpose:
    Gather data from Citrix licensing server using WMI. Data is output to CSV file and then utilized to create an ASP.NET graph and a web page.
    Script can gather CCU (concurrent) and U/D (user device) data.
	Script could be edited to run against multiple license servers and aggregate the data.

    ***Pre-requisities***
 	- Install Microsoft Chart Controls for .NET - https://www.microsoft.com/en-us/download/details.aspx?id=14422
      - This is required for the ASP.NET charting within the script and should be installed on the hosting server
    - Edit the JSON file using the template as a base

    Versions:
    3/25/2018 - Version 1.00 - COMPLETE - Alex Spicola

#>
#=========================================================================================================================
#Set script run directory
$runDir = $PSScriptRoot

#Start transcript
Start-Transcript "$runDir\log.log"

#Import the JSON file
#$JSONInput = ConvertFrom-Json "$(Get-Content "./licenseusage.json")" #Used for dev/test
$JSONInput = ConvertFrom-Json "$(Get-Content "$runDir\licenseusage.json")"

#Convert JSON inputs to varibles for use in script
$scriptDir = $JSONInput.rundirectory #Directory the script runs
$webDir = $JSONInput.webdirectory #Web directory to place any web files
[int]$sleepMin = $JSONInput.scriptsleep #Minutes for cycle sleep time, use in infinite loop to sleep the script for a period of time
$maxMonths = $JSONInput.maxmonths #Sets limit for number of months to use for data publishing (stats and graphs)
$file = $JSONInput.outputfilename #Output file name for license data CSV, data will append
$licSrv = $JSONInput.licenserver
$prod = $JSONInput.product
$prodedition = $JSONInput.prodedition
$lictype = $JSONInput.lictype
$loop = $JSONInput.loop #Infinifite loop variable

#Script output folder location, CSV file name, web output location
$outfolder = "$($scriptDir)"
$FileName = $($outfolder)+$($file) #Combine output folder and filename

#=========================================================================================================================
Write-Output "Start the script"
Write-Output "Output location and file: $FileName"

#Start the loop, this will only continue configured in JSON
do {

#=========================================================================================================================
#Gather all licensing data and output to usable CSV
#Original code sources, used and edited some pieces.
#https://www.jonathanmedd.net/2011/01/monitor-citrix-license-usage-with-powershell.html
#Original code from previous Citrix administrator edited to needs
#=========================================================================================================================
#Variables
Write-Output "Create variables"
$Total = 0
$InUse = 0

#Date and time for CSV output
Write-Output "Set date and time for CSV output"
$Date = Get-Date -format "MM/dd/yyyy" ; $Time = Get-Date -format "HH:mm tt"

#Output license server names
Write-Output "Citrix license server: $($licSrv)"

#Get Citrix licensing Info
Write-Output "Set WMI call for license server"
$licPool = gwmi -class "Citrix_GT_License_Pool" -Namespace "Root\CitrixLicensing" -comp $licSrv
	
#Gather license data for CCU (concurrent user) licenses, ensure to update the for proper version and license type for WMI object
Write-Output "Gather license data from license server"
#Get product from JSON information
if ( $prod -eq "xa" ) #XenApp
{ 
	$prd = "MPS" 
} 
elseif ( $prod -eq "xd" ) #XenDesktop
{ 
	$prd = "XDT" 
}

#Get product edition from JSON information
if ( $prodedition -eq "ent" ) #Enterprise licensing
{ 
$ed = "ENT" 
} 
elseif ( $prodedition -eq "plt" ) #Platinum licensing
{
	$ed = "PLT" 
} 
	
#Get type from JSON information
if ( $lictype -eq "ccu" ) #Concurrent user
{ 
	$type = "CCU" 
} 
elseif ( $lictype -eq "ud" ) #User/Device
{ 
	$type = "UD" 
} 

$licPool | ForEach-Object{ If ($_.PLD -eq "$($prd)_$($ed)_$($type)") {
    $Total = $Total + $_.Count
    $InUse = $InUse + $_.InUseCount
    }
}

#Calculate the totals and percentages with data just pulled
Write-Output "Calculate percentages"
$PctUsed = [Math]::Round($InUse/$Total*100,0)
$Free = [Math]::Round($Total-$InUse)

#Create hashtable object and export the data
Write-Output "Create object and add data"
$obj = New-Object psobject
$obj | Add-Member -MemberType NoteProperty -Name Date -Value $Date
$obj | Add-Member -MemberType NoteProperty -Name Time -Value $Time
$obj | Add-Member -MemberType NoteProperty -Name Total -Value $Total
$obj | Add-Member -MemberType NoteProperty -Name InUse -Value $InUse
$obj | Add-Member -MemberType NoteProperty -Name Free -Value $Free
$obj | Add-Member -MemberType NoteProperty -Name PctUsed -Value $PctUsed

#Export and append the data to a CSV file
Write-Output "Output data to CSV"
$obj | Export-Csv $FileName -Append

#=========================================================================================================================
#Data Import
#Import data from CSV file and cutback to month limit ($dataLimit variable)
#=========================================================================================================================
Write-Output "Set start date for data"
$startDate = (Get-Date).AddMonths(-$maxMonths)
$startDateFmt = ($startDate | Get-Date -Format "MM/dd/yyyy")

Write-Output "Maximum months to import: $($maxMonths) months"

Write-Output "Import data from $($FileName)"
$licData = Import-Csv $FileName | ? { [datetime]$_.Date -ge $startDateFmt }
#=========================================================================================================================
#Graph Creation
#Code sources:
#https://goodworkaround.com/2014/06/18/graphing-with-powershell-done-easy/
#https://learn-powershell.net/2016/09/18/building-a-chart-using-powershell-and-chart-controls/
#https://blogs.technet.microsoft.com/richard_macdonald/2009/04/28/charting-with-powershell/
#https://gallery.technet.microsoft.com/scriptcenter/Charting-Line-Chart-using-df47af9c
#=========================================================================================================================
Write-Output "Create graphs from data"

#Take license data and cut down to maximum per day for graphing
Write-Output "Create variables and arrays"
#Variables
$licMaxDataAll = @()

#Create array of unique dates to compare against
Write-Output "Get unique dates from data"
$dates = $licData | select date -Unique

#Check each individual date and find the maximum license in use count for all entries on that day
#Output those entries to new array, cut down to single entry if there are multiple, write single entry to object for graphing
Write-Output "Find maximum for each date"
foreach ( $d in $dates ) 
{	
	$dayArr = @()

    foreach ( $entry in $licData ) { if ($entry.Date -eq $d.Date) { $dayArr += $entry } }
    
    $max = ( $dayArr | measure -Property InUse -Maximum ).Maximum

    $daymax = @( $dayArr | ? { $_.InUse -eq $max } )

    if ( $daymax.count -gt 1 ) { $daymax = $daymax[0] } #If multiple duplicate max entries, cut down to first one only

    $licMaxDataAll += $daymax   
}

Write-Output "Total entries in data file: $($licMaxDataAll.Count)"

#Create realtime data points (last 3 hours) from all license data
#12 data points per hour (5 minute runs), totals 36 entries for last 3 hours
Write-Output "Create realtime datapoints, last 3 hours of current date"
$licDataRT = $licData | select -Last 36
$licMaxData = $licMaxDataAll

#Get maximum and average for current month
Write-Output "Set maximum and average for current month"
$splitDate = ($date -split "/")
$day = $splitDate[1]
$maxCurrMonth = ($licMaxData | select -Last $day | measure -Property InUse -Maximum).Maximum
$avgCurrMonth = [math]::round(($licMaxData | select -Last $day | measure -Property InUse -Average).Average)

#Get maximums and averages for last 30 days or all data if less than 30 days
Write-Output "Set maximum and averages for last 30 days"
$max30days = ($licMaxData | select -Last 30 | measure -Property InUse -Maximum).Maximum
$last30Avg = [math]::round(($licMaxData | select -Last 30 | measure -Property InUse -Average).Average) #Average for last 30 days
$last30AvgPct = [math]::round(($licMaxData | select -Last 30 | measure -Property PctUsed -Average).Average) #Average percentage for last 30 days
$last30Max  = $licMaxData | select -Last 30 | ? { $_.InUse -eq $max30days } #Max for last 30 days

#Chart creation options, used to create different charts for realtime (last 3 hours), last 7 days, last 30 days, and last 180 days
#These are used for different link images in the dashboard
Write-Output "Set chart options"
$chartOpts = @("RT",7,30,"MaxMonths") #Chart options are RT (realtime), 7 days, 30 days, and the maximum months from the JSON file

$chartOpts | % {
    Write-Output "Current chart option: $($_)"

    if ($_ -eq "RT") { $chartData = $licDataRT }
    if ($_ -eq "7") { $chartData = $licMaxData | select -Last 7 | sort -Descending }
    if ($_ -eq "30") { $chartData = $licMaxData | select -Last 30 | sort -Descending }
    if ($_ -eq "MaxMonths") { $chartData = $licMaxData }

    #Set the chart interval based on data count
    if ($_ -eq "RT") { $interval = 1 }
    if ($_ -eq "7") { $interval = 1 }
    if ($_ -eq "30") { $interval = 2 }
    if ($_ -eq "MaxMonths") { $interval = 5 }
    Write-Output "Chart interval: $($interval)"

    #Load charting controls
    Write-Output "Load chart controls"
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    #Create chart base
    Write-Output "Create chart base for $($_) chart"
    $LicUsageChart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $LicUsageChart.Width = 1200
    $LicUsageChart.Height = 800
    $LicUsageChart.BackColor = [System.Drawing.Color]::White

    #Create chart area
    Write-Output "Create chart area for $($_) chart"
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea"
    $LicUsageChart.ChartAreas.Add($chartarea)

    #Set chart styles
    Write-Output "Set chart styles for $($_) chart"
    $chartarea.AxisX.IsLabelAutoFit = $true
    $chartarea.AxisX.LabelStyle.Angle = "75"
    $chartarea.AxisX.Interval = $interval

    #Create chart titles
	if ($_ -eq "RT") { [void]$LicUsageChart.Titles.Add("Realtime (last 3 hours)") }
	if ($_ -eq "7") { [void]$LicUsageChart.Titles.Add("Last 7 Days") }
	if ($_ -eq "30") { [void]$LicUsageChart.Titles.Add("Last 30 Days") }
	if ($_ -eq "MaxMonths") { [void]$LicUsageChart.Titles.Add("Last $($maxMonths) Months") }
		

    #All in use licenses chart series
    Write-Output "Add all in use data series for $($_) chart"
    [void]$LicUsageChart.Series.Add("All In Use")
    $LicUsageChart.Series["All In Use"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    if ($_ -eq "RT") { $LicUsageChart.Series["All In Use"].Points.DataBindXY($chartData.Time, $chartData.InUse) } else { $LicUsageChart.Series["All In Use"].Points.DataBindXY($chartData.Date, $chartData.InUse) }
    $LicUsageChart.Series["All In Use"].Color = "Green"
    $LicUsageChart.Series["All In Use"].BorderWidth = 5

    #Total owned licenses chart series
    Write-Output "add total owned data series for $($_) chart"
    [void]$LicUsageChart.Series.Add("Licenses Owned")
    $LicUsageChart.Series["Licenses Owned"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $LicUsageChart.Series["Licenses Owned"].Points.DataBindXY($chartData.Date, $chartData.AllTotal)
    $LicUsageChart.Series["Licenses Owned"].Color = "Red"
    $LicUsageChart.Series["Licenses Owned"].BorderWidth = 3

    #Trendline of all in use chart series (if not realtime)
    if ($_ -ne "RT") 
    {
        Write-Output "Add trendline data series for $($_) chart"
        [void]$LicUsageChart.Series.Add("Trend")
        $LicUsageChart.Series["Trend"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
        $LicUsageChart.Series["Trend"].BorderDashStyle = "Dot"
        $LicUsageChart.Series["Trend"].Color = "Black"
        $LicUsageChart.Series["Trend"].BorderWidth = 3
        $LicUsageChart.DataManipulator.FinancialFormula("Forecasting", "Linear,0,false,false", $LicUsageChart.Series["All In Use"], $LicUsageChart.Series["Trend"])
    }

    #Create chart legend
    Write-Output "Create chart legend for $($_) chart"
    [void]$LicUsageChart.Legends.Add("Legend")
    $LicUsageChart.Legends["Legend"].Font = "segoeuilight,10pt"
    $LicUsageChart.Legends["Legend"].Docking = "Bottom"
    $LicUsageChart.Legends["Legend"].Alignment = "Center"

    <#Show the chart - for testing purposes
    $LicUsageChart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left 
    $Form = New-Object Windows.Forms.Form 
    $Form.Width = 1500
    $Form.Height = 725
    $Form.controls.add($LicUsageChart) 
    $Form.Add_Shown({$Form.Activate()}) 
    $Form.ShowDialog()#>

    #Output the charts as PNG  files and copy to web server
    Write-Output "Create PNG file for $($_) chart and copy to web folder"
    if ( Test-Path "$($outfolder)CtxLicUsage$($_).png") { Remove-Item "$($outfolder)\CtxLicUsage$($_).png"} 
    $LicUsageChart.SaveImage("$($outfolder)\CtxLicUsage$($_).png","png")
    Copy-Item "$($outfolder)\CtxLicUsage$($_).png" "$($webdir)"
}

#=========================================================================================================================
#Website creation
#=========================================================================================================================
Write-Output "Create webpage"

#Varibles
Write-Output "Set variables"
$title = "Citrix License Usage"
$webdate = ( Get-Date -format g )
$currMonth = (Get-Culture).DateTimeFormat.GetMonthName(((Get-Date).Month))

#Create web header
Write-Output "Create web header"
$head = @"
<html>
<head>
<title>$title</title>
"@
$style = @"
<style>
th {
	font-family: Tahoma;
	font-size: 14px;
	padding-top: 1px;
	padding-right: 1px;
	padding-bottom: 1px;
	padding-left: 1px;
	overflow: hidden;
}

td {
   font-family: Tahoma;
	font-size: 14px;
	padding-top: 1px;
	padding-right: 1px;
	padding-bottom: 1px;
	padding-left: 1px;
	overflow: hidden;
}
.liccontent a {
    color:blue;
}
.liccontent a:active {
    color:green;
}
.liccontent a:visited {
    color:blue;
}
</style>
<script src="https://code.jquery.com/jquery-3.3.1.js"></script>
<script>
    function changeImage(element) {
	    document.getElementById('imageReplace').src = element;
    }
</script
</head>
"@

#Create web body and add licensing data
Write-Output "Create web body and add all charts and data"
$body = @"
<body>
<div class="liccontent">
<center>
<font face='tahoma' color='#000000' size='5'><strong>$title</strong></font>
<p>
<table border="0" width="100%">
    <tr align="left" border="0">
        <td align="right" valign="top" border="0" width="20%">
            <table border="0">
                <tr>
                    <th colspan=3><font size="3">Current</font><hr></th>
                </tr>
                <tr>
                    <th>In Use</th><td width="10"></td><td>$InUse</td>
                </tr>
                <tr>
                    <th>% In Use</th><td></td><td>$PctUsed%</td>
                </tr>
                <tr>
                    <th>Plt In Use</th><td></td><td>$PlatInUse</td>
                </tr>
                <tr>
                    <th>Ent In Use</th><td></td><td>$EntInUse</td>
                </tr>
                <tr>
                    <th>Lic. Owned</th><td></td><td>$Total</td>
                </tr>
                <tr>
                    <th>Ent. Lic. Reserve</th><td></td><td>$EntTotal</td>
                </tr>
                <tr height="14"><td></td></tr>
                <tr>
                    <th>$($currMonth) $($splitDate[2]) Max</th><td></td><td><b>$maxCurrMonth</b></td>
                </tr>
                <tr>
                    <th>$($currMonth) $($splitDate[2]) Avg</th><td></td><td>$avgCurrMonth</td>
                </tr>
                <tr height="14"><td></td></tr>
                <tr>
                    <th colspan=3><font size="3">Historical</font><hr></th>
                </tr>
                <tr>
                    <th>30 Day Avg</th><td></td><td>$last30Avg</td>
                </tr>
                <tr>
                    <th>30 Day Avg %</th><td></td><td>$last30AvgPct%</td>
                </tr>
                <tr>
                    <th>30 Day Max</th><td></td><td>$($last30Max.InUse) ($($last30Max.Date))</td>
                </tr>
                <tr>
                    <th>Plt 30 Day Max</th><td></td><td>$($last30PltMax.PlatInUse) ($($last30PltMax.Date))</td>
                </tr>
                <tr>
                    <th>Ent 30 Day Max</th><td></td><td>$($last30EntMax.EntInUse) ($($last30EntMax.Date))</td>
                </tr>
            </table>
        </td>
        <td align="center">
        <a href="#" onclick="changeImage('./CtxLicUsageRT.png');">Realtime</a> | <a href="#" onclick="changeImage('CtxLicUsage7.png');">Last 7 Days</a> | <a href="#" onclick="changeImage('./CtxLicUsage30.png');">Last 30 Days</a> | <a href="#" onclick="changeImage('./CtxLicUsageMaxMonths.png');">Last $($maxMonths) Months</a>
	    <br>
		<img src = "./CtxLicUsageRT.png" id="imageReplace"/>
	    </td>
    </tr>
</table>
<br>
<font size="3"><b>Last Updated:</b> $webdate CST</font>
</div>
</center>
</body>
</html>
"@

Write-Output "Output all web data"
$head | Out-File "$webDir\ctxlicensing.htm"
$style | Out-File "$webDir\ctxlicensing.htm" -Append
$body | Out-File "$webDir\ctxlicensing.htm" -Append

if ( $loop -eq "true" ) { 
    #Sleep between script loop cycles
    Write-Output "Script will sleep for $($sleepMin) minutes" 
    $sleeptimer = $sleepMin * 60 #Converts sleep minutes to seconds
    Start-Sleep $sleeptimer
}
	
	
}
while ($loop -eq "true")

Stop-Transcript

