import-module .\OutlookTools.psm1

$shabbatTimes = "5 Jan 2018 	3:50pm
12 Jan 2018 	4:01pm
19 Jan 2018 	4:12pm
26 Jan 2018 	4:24pm
2 Feb 2018 	4:36pm 
9 Feb 2018 	4:49pm
16 Feb 2018 	5:02pm
23 Feb 2018 	5:15pm
2 Mar 2018 	5:28pm
9 Mar 2018 	5:40pm
16 Mar 2018 	5:52pm
23 Mar 2018 	6:04pm" -Split "\n"

$minutesRequiredBeforeShabbat = 120

foreach ($shabbatStartTimeString in $shabbatTimes)
{
    #$shabbatStartTime = [datetime]::Parse("5 Jan 2018 3:50pm")
    $shabbatStartTime = [datetime]::Parse($shabbatStartTimeString)
    $timeRequiredBeforeShabbat = [timespan]::FromMinutes($minutesRequiredBeforeShabbat)

    $meetinStart = $shabbatStartTime - $timeRequiredBeforeShabbat

    echo "Create meeting from ${meetinStart} for duration of 240 min"
    New-OutlookCalendarMeeting -Subject Shabbat -BusyStatus 3 -Body " " -Location " " -MeetingStart $meetinStart -MeetingDuration 240
    
}