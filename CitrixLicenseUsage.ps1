#=========================================================================================================================
<#
    Author: Alex Spicola
    Created: March 2018
    
    Purpose:
    Gather data from Citrix licensing server using WMI. Data is output to CSV file and then utilized to create an ASP.NET graph and a web page.
    Script can gather CCU (concurrent) and U/D (user device) data.
	Script could be edited to run against multiple license servers and aggregate the data.
    No charts will begin to be created until enough datapoints exist, this means greater than one single run. For more than a real-time chart multiple days of data are required.
    Charts will create starting with >1 datapoint but they will not be the for the timeframe list. This will fill in over time as the data is gathered. This will require a continual loop or scheduled task.

    ***Pre-requisities***
 	- Install Microsoft Chart Controls for .NET - https://www.microsoft.com/en-us/download/details.aspx?id=14422
      - This is required for the ASP.NET charting within the script and should be installed on the hosting server
    - Edit the JSON file using the template as a base

    Versions:
    3/25/2018 - Version 1.00
	3/27/2018 - Version 2.00
		- Added CCS (Citrix Customer Select) to license type
    3/28/2018 - Version 3.00 - CURRENT VERSION
        - Removed additional hardcoding from original
        - Updated outputs
        - Create web directory as necessary
        - Updated variables to PascalCase
        - Removed some whitespace
        - Changed logfile name
        - Updated chart creation
            - Single data point will not create any charts, multiple datapoints are required
            - Less than a day of data will create only real-time chart
            - Greater than a day of data will create all charts
        - Added HTML change to output text to page when no charts are created
#>
#=========================================================================================================================
#Set script run directory
$RunDir = $PSScriptRoot

#Start transcript
Start-Transcript "$RunDir\ctxlicusglog.log"

#Date and time for CSV output
Write-Output "$($Date):  Set date and time for CSV output"
$Date = Get-Date -format "MM/dd/yyyy" ; $Time = Get-Date -format "HH:mm tt"

#Import the JSON file
$JSONInput = ConvertFrom-Json "$(Get-Content "$runDir\licenseusage.json")"
#$JSONInput = ConvertFrom-Json "$(Get-Content ".\licenseusage.json")"#Used for testing

#Convert JSON inputs to varibles for use in script
$ScriptDir = $JSONInput.rundirectory #Directory the script runs
$WebDir = $JSONInput.webdirectory #Web directory to place any web files
[int]$SleepMin = $JSONInput.scriptsleep #Minutes for cycle sleep time, use in infinite loop to sleep the script for a period of time
$MaxMonths = $JSONInput.maxmonths #Sets limit for number of months to use for data publishing (stats and graphs)
$File = $JSONInput.outputfilename #Output file name for license data CSV, data will append
$LicSrv = $JSONInput.licenserver
$Prod = $JSONInput.product
$ProdEdition = $JSONInput.prodedition
$LicType = $JSONInput.lictype
$Loop = $JSONInput.loop #Infinifite loop variable

#Set web directory/out folder location and CSV file name
if ($WebDir -ne $ScriptDir) { if (!(Test-Path $WebDir)) {New-Item -ItemType directory -Path $WebDir | Out-Null}} #Create ouput/webdirectory if necessary
$FileName = $($WebDir)+$($File) #Combine output folder and filename

#=========================================================================================================================
Write-Output "$($Date):  $($Date)Start the script"
Write-Output "$($Date):  $($Date) Output location and file: $FileName"

#Start the loop, this will only continue if configured in JSON
do {

#=========================================================================================================================
#Gather all licensing data and output to usable CSV
#Original code sources, used and edited some pieces.
#https://www.jonathanmedd.net/2011/01/monitor-citrix-license-usage-with-powershell.html
#Original code from previous Citrix administrator edited to needs
#=========================================================================================================================
#Variables
Write-Output "$($Date):  $($Date) Create variables"
$Total = 0
$InUse = 0

#Output license server names
Write-Output "$($Date):  Citrix license server: $($LicSrv)"

#Get Citrix licensing Info
Write-Output "$($Date):  Set WMI call for license server"
$LicPool = gwmi -class "Citrix_GT_License_Pool" -Namespace "Root\CitrixLicensing" -comp $LicSrv
if ($LicPool -ne $null) {Write-Output "$($Date):  WMI pull was successful"} else { Write-Output "$($Date):  WMI pull was NOT successful"}
	
#Gather license data for CCU (concurrent user) licenses, ensure to update the for proper version and license type for WMI object
Write-Output "$($Date):  Gather license data from license server"
#Get product from JSON information
if ($Prod -eq "xa") #XenApp
{ 
	$Prd = "MPS" 
} 
elseif ($Prod -eq "xd" -or $Prod -eq "xdt") #XenDesktop
{ 
	$Prd = "XDT" 
}

#Get product edition from JSON information
if ($ProdEdition -eq "ent") #Enterprise licensing
{ 
    $Ed = "ENT" 
} 
elseif ($ProdEdition -eq "plt") #Platinum licensing
{
	$Ed = "PLT" 
} 
	
#Get type from JSON information
if ($LicType -eq "ccu") #Concurrent user
{ 
	$Type = "CCU" 
} 
elseif ($LicType -eq "ud") #User/Device
{ 
	$Type = "UD" 
} 
elseif ($LicType -eq "ccs") #Citrix Customer Select
{
	$Type ="CCS"
}

Write-Output "$($Date):  License information: $($Prd)_$($Ed)_$($Type)"

#Gather totals and in use count
Write-Output "$($Date):  Calculate percentages"
$LicPool | ForEach-Object{ If ($_.PLD -eq "$($Prd)_$($Ed)_$($Type)") {
    $Total = $Total + $_.Count
    $InUse = $InUse + $_.InUseCount
    }
}
Write-Output "$($Date):  Calculated - Total: $($Total), In Use $($InUse)"

#Calculate the totals and percentages with data just pulled
Write-Output "$($Date):  Calculate percentages"
$PctUsed = [Math]::Round($InUse/$Total*100,0)
$Free = [Math]::Round($Total-$InUse)

#Create hashtable object and export the data
Write-Output "$($Date):  Create object and add data"
$Obj = New-Object psobject
$Obj | Add-Member -MemberType NoteProperty -Name Date -Value $Date
$Obj | Add-Member -MemberType NoteProperty -Name Time -Value $Time
$Obj | Add-Member -MemberType NoteProperty -Name Total -Value $Total
$Obj | Add-Member -MemberType NoteProperty -Name InUse -Value $InUse
$Obj | Add-Member -MemberType NoteProperty -Name Free -Value $Free
$Obj | Add-Member -MemberType NoteProperty -Name PctUsed -Value $PctUsed

#Export and append the data to a CSV file
Write-Output "$($Date):  Output data to CSV"
$Obj | Export-Csv $FileName -Append

#=========================================================================================================================
#Data Import
#Import data from CSV file and cutback to month limit ($dataLimit variable)
#=========================================================================================================================
Write-Output "$($Date):  Set start date for data"
$StartDate = (Get-Date).AddMonths(-$MaxMonths)
$StartDateFmt = ($StartDate | Get-Date -Format "MM/dd/yyyy")

Write-Output "$($Date):  Maximum months to import: $($MaxMonths) months"

Write-Output "$($Date):  Import data from $($FileName)"
$LicData = @(Import-Csv $FileName | ? {[datetime]$_.Date -ge $StartDateFmt})
#=========================================================================================================================
#Graph Creation
#Code sources:
#https://goodworkaround.com/2014/06/18/graphing-with-powershell-done-easy/
#https://learn-powershell.net/2016/09/18/building-a-chart-using-powershell-and-chart-controls/
#https://blogs.technet.microsoft.com/richard_macdonald/2009/04/28/charting-with-powershell/
#https://gallery.technet.microsoft.com/scriptcenter/Charting-Line-Chart-using-df47af9c
#=========================================================================================================================
Write-Output "$($Date):  Create graphs from data"

#Take license data and cut down to maximum per day for graphing
Write-Output "$($Date):  Create variables and arrays"
#Variables
$LicMaxDataAll = @()

#Create array of unique dates to compare against
Write-Output "$($Date):  Get unique dates from data"
$Dates = $LicData | select date -Unique

#Check each individual date and find the maximum license in use count for all entries on that day
#Output those entries to new array, cut down to single entry if there are multiple, write single entry to object for graphing
Write-Output "$($Date):  Find maximum for each date"
foreach ($D in $Dates) 
{	
	$DayArr = @()

    foreach ($Entry in $licData) {if ($Entry.Date -eq $D.Date) {$DayArr += $Entry}}
    
    $Max = ($DayArr | measure -Property InUse -Maximum).Maximum

    $DayMax = @($DayArr | ? {$_.InUse -eq $Max})
        
    if ($DayMax.count -gt 1) {$DayMax = $DayMax[0]} #If multiple duplicate max entries, cut down to first one only

    $LicMaxDataAll += $DayMax   
}

Write-Output "$($Date):  Total entries in data file: $($LicMaxDataAll.Count)"

#Create real-time data points from all license data
Write-Output "$($Date):  Create real time datapoints"
$LicDataRT = $LicData | select -Last 40
$LicMaxData = $LicMaxDataAll

#Get maximum and average for current month
Write-Output "$($Date):  Set maximum and average for current month"
$SplitDate = ($Date -split "/")
$Day = $SplitDate[1]
$MaxCurrMonth = ($LicMaxData | select -Last $Day | measure -Property InUse -Maximum).Maximum
$AvgCurrMonth = [math]::round(($LicMaxData | select -Last $Day | measure -Property InUse -Average).Average)

#Get maximums and averages for last 30 days or all data if less than 30 days
Write-Output "$($Date):  Set maximum and averages for last 30 days"
$Max30Days = ($LicMaxData | select -Last 30 | measure -Property InUse -Maximum).Maximum
$Last30Avg = [math]::round(($LicMaxData | select -Last 30 | measure -Property InUse -Average).Average) #Average for last 30 days
$Last30AvgPct = [math]::round(($LicMaxData | select -Last 30 | measure -Property PctUsed -Average).Average) #Average percentage for last 30 days
$Last30Max  = $LicMaxData | select -Last 30 | ? { $_.InUse -eq $Max30Days } #Max for last 30 days

#Chart creation options, used to create different charts for real-time (last 3 hours), last 7 days, last 30 days, and last 180 days
Write-Output "$($Date):  Set chart options based on datapoints"
if ($LicMaxData.Count -eq 1) {$ChartOpts = $null} #Charts will not be created if only a single datapoint is found
if ($LicMaxData.Count -gt 1) {$ChartOpts = @("RT",7,30,"MaxMonths")} #Chart options set to RT (real-time), 7 days, 30 days, and the maximum months from the JSON file.

#Will only create charts when more than one datapoint is found
if ($ChartOpts -ne $null) 
{
Write-Output "$($Date):  Minimum datapoints found, start chart creation"
    $ChartOpts | % {
        Write-Output "$($Date):  Current chart option: $($_)"

        if ($_ -eq "RT") {$ChartData = $LicDataRT}
        if ($_ -eq "7") {$ChartData = $LicMaxData | select -Last 7 | sort -Descending}
        if ($_ -eq "30") {$ChartData = $LicMaxData | select -Last 30 | sort -Descending}
        if ($_ -eq "MaxMonths") {$ChartData = $LicMaxData}

        #Set the chart interval based on data count
        if ($_ -eq "RT") {$Interval = 1}
        if ($_ -eq "7") {$Interval = 1}
        if ($_ -eq "30") {$Interval = 2}
        if ($_ -eq "MaxMonths") {$Interval = 5}
        Write-Output "$($Date):  Chart interval: $($Interval)"

        #Load charting controls
        Write-Output "$($Date):  Load chart controls"
        [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

        #Create chart base
        Write-Output "$($Date):  Create chart base for $($_) chart"
        $LicUsageChart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
        $LicUsageChart.Width = 1200
        $LicUsageChart.Height = 800
        $LicUsageChart.BackColor = [System.Drawing.Color]::White

        #Create chart area
        Write-Output "$($Date):  Create chart area for $($_) chart"
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $ChartArea.Name = "ChartArea"
        $LicUsageChart.ChartAreas.Add($ChartArea)

        #Set chart styles
        Write-Output "$($Date):  Set chart styles for $($_) chart"
        $ChartArea.AxisX.IsLabelAutoFit = $True
        $ChartArea.AxisX.LabelStyle.Angle = "75"
        $ChartArea.AxisX.Interval = $Interval

        #Create chart titles
	    if ($_ -eq "RT") {[void]$LicUsageChart.Titles.Add("Real-Time")}
	    if ($_ -eq "7") {[void]$LicUsageChart.Titles.Add("Last 7 Days")}
	    if ($_ -eq "30") {[void]$LicUsageChart.Titles.Add("Last 30 Days")}
	    if ($_ -eq "MaxMonths") {[void]$LicUsageChart.Titles.Add("Last $($MaxMonths) Months")}
		
        #In Use licenses chart series
        Write-Output "$($Date):  Add in use data series for $($_) chart"
        [void]$LicUsageChart.Series.Add("In Use")
        $LicUsageChart.Series["In Use"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
        if ($_ -eq "RT") {$LicUsageChart.Series["In Use"].Points.DataBindXY($ChartData.Time, $ChartData.InUse)} 
        else {$LicUsageChart.Series["In Use"].Points.DataBindXY($ChartData.Date, $ChartData.InUse)}
        $LicUsageChart.Series["In Use"].Color = "Green"
        $LicUsageChart.Series["In Use"].BorderWidth = 5

        #Licenses Owned licenses chart series
        Write-Output "$($Date):  Add total owned data series for $($_) chart"
        [void]$LicUsageChart.Series.Add("Licenses Owned")
        $LicUsageChart.Series["Licenses Owned"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
        $LicUsageChart.Series["Licenses Owned"].Points.DataBindXY($ChartData.Date, $ChartData.Total)
        $LicUsageChart.Series["Licenses Owned"].Color = "Red"
        $LicUsageChart.Series["Licenses Owned"].BorderWidth = 3

        #Trendline of In Use chart series (if not real time)
        if ($_ -ne "RT") 
        {
            Write-Output "$($Date):  Add trendline data series for $($_) chart"
            [void]$LicUsageChart.Series.Add("Trend")
            $LicUsageChart.Series["Trend"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
            $LicUsageChart.Series["Trend"].BorderDashStyle = "Dot"
            $LicUsageChart.Series["Trend"].Color = "Black"
            $LicUsageChart.Series["Trend"].BorderWidth = 3
            $LicUsageChart.DataManipulator.FinancialFormula("Forecasting", "Linear,0,false,false", $LicUsageChart.Series["In Use"], $LicUsageChart.Series["Trend"])
        }

        #Create chart legend
        Write-Output "$($Date):  Create chart legend for $($_) chart"
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
        Write-Output "$($Date):  Create PNG file for $($_) chart and copy to web folder"
        if ( Test-Path "$($WebDir)CtxLicUsage$($_).png") { Remove-Item "$($WebDir)\CtxLicUsage$($_).png"} 
        $LicUsageChart.SaveImage("$($WebDir)\CtxLicUsage$($_).png","png")
    }
} else {Write-Output "$($Date):  Minimum datapoints not found, no charts created"}

#=========================================================================================================================
#Website creation
#=========================================================================================================================
Write-Output "$($Date):  Create webpage"

#Varibles
Write-Output "$($Date):  Set variables"
$Title = "Citrix License Usage"
$WebDate = (Get-Date -format g)
$CurrMonth = (Get-Culture).DateTimeFormat.GetMonthName(((Get-Date).Month))

#Set HTML for links or no data based on chart options (datapoints)
if ($ChartOpts -ne $null) {
$ChartLinks = @"
    <a href="#" onclick="changeImage('./CtxLicUsageRT.png');">Real-Time</a> | <a href="#" onclick="changeImage('CtxLicUsage7.png');">Last 7 Days</a> | <a href="#" onclick="changeImage('./CtxLicUsage30.png');">Last 30 Days</a> | <a href="#" onclick="changeImage('./CtxLicUsageMaxMonths.png');">Last $($maxMonths) Months</a>
    <br>
    <img src = "./CtxLicUsageRT.png" id="imageReplace"/>
"@
} else {
$ChartLinks = @"
<h3>NO CHART DATA AVAILABLE AT THIS TIME</h3>
"@
}

#Create web header
Write-Output "$($Date):  Create web header"
$Head = @"
<html>
<head>
<title>$Title</title>
"@
$Style = @"
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
Write-Output "$($Date):  Create web body and add all charts and data"
$Body = @"
<body>
<div class="liccontent">
<center>
<font face='tahoma' color='#000000' size='5'><strong>$Title</strong></font>
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
                <tr>
                    <th>Lic. Owned</th><td></td><td>$Total</td>
                </tr>
                <tr height="14"><td></td></tr>
                <tr>
                    <th>$($CurrMonth) $($SplitDate[2]) Max</th><td></td><td><b>$MaxCurrMonth</b></td>
                </tr>
                <tr>
                    <th>$($CurrMonth) $($SplitDate[2]) Avg</th><td></td><td>$AvgCurrMonth</td>
                </tr>
                <tr height="14"><td></td></tr>
                <tr>
                    <th colspan=3><font size="3">Historical</font><hr></th>
                </tr>
                <tr>
                    <th>30 Day Avg</th><td></td><td>$Last30Avg</td>
                </tr>
                <tr>
                    <th>30 Day Avg %</th><td></td><td>$Last30AvgPct%</td>
                </tr>
                <tr>
                    <th>30 Day Max</th><td></td><td>$($Last30Max.InUse) ($($Last30Max.Date))</td>
                </tr>
            </table>
        </td>
        <td align="center">
        $ChartLinks
	    </td>
    </tr>
</table>
<br>
<font size="3"><b>Last Updated:</b> $WebDate</font>
</div>
</center>
</body>
</html>
"@

Write-Output "$($Date):  Output all web data"
$Head | Out-File "$WebDir\ctxlicensing.htm"
$Style | Out-File "$WebDir\ctxlicensing.htm" -Append
$Body | Out-File "$WebDir\ctxlicensing.htm" -Append

if ( $Loop -eq "true" ) { 
    #Sleep between script loop cycles
    Write-Output "$($Date):  Script will sleep for $($SleepMin) minutes" 
    $SleepTimer = $SleepMin * 60 #Converts sleep minutes to seconds
    Start-Sleep $SleepTimer
}
	
} while ($Loop -eq "true")

Stop-Transcript