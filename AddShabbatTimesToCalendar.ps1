import-module .\OutlookTools.psm1

# Config
$city = 'GB-London'
$minutesRequiredBeforeShabbat = 120
$year = 2020
$months = 9, 10, 11, 12

# Code
foreach ($month in $months)
{
    $url = "https://www.hebcal.com/hebcal?v=1&cfg=json&c=on&m=0&year=$year&month=$month&city=$city"

    $shabbatTimes = ((curl $url).Content | ConvertFrom-Json).items | select-object  date 

    foreach ($ShabbatStartTimeObj in $shabbatTimes)
    {
        $shabbatStartTime = [datetime]::Parse($ShabbatStartTimeObj.date)
        $timeRequiredBeforeShabbat = [timespan]::FromMinutes($minutesRequiredBeforeShabbat)

        $meetinStart = $shabbatStartTime - $timeRequiredBeforeShabbat

        echo "Create meeting from ${meetinStart} for duration of 240 min"
        New-OutlookCalendarMeeting -Subject Shabbat -BusyStatus 3 -Body " " -Location " " -MeetingStart $meetinStart -MeetingDuration 240
        
    }
}