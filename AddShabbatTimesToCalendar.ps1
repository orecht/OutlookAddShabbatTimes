import-module .\OutlookTools.psm1
param (
    ## Location 
    [Parameter(Mandatory=$false, HelpMessage="City code as defined at https://github.com/hebcal/dotcom/blob/master/hebcal.com/dist/cities2.txt")]
    $city = 'GB-London',

    ## Number of minutes required before shabbat (travel home time + preparation)
    [Parameter(Mandatory=$false, HelpMessage="Number of minutes required before shabbat (travel home time + preparation) (default 60 min)")]
    $minutesRequiredBeforeShabbat = 60,

    # Dates to insert
    [Parameter(Mandatory=$false, HelpMessage="Year")]
    $year = (Get-Date).Year,
    [Parameter(Mandatory=$false, HelpMessage="Months")]
    $months = (Get-Date -Format "MM")
)

function Show-Usage {
    Write-Host "Usage: AddShabbatTimesToCalendar.ps1 [-city <city>] [-minutesRequiredBeforeShabbat <minutes>] [-year <year>] [-months <months>]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -city                      City code as defined at https://github.com/hebcal/dotcom/blob/master/hebcal.com/dist/cities2.txt"
    Write-Host "  -minutesRequiredBeforeShabbat   Number of minutes required before shabbat (travel home time + preparation) (default 60 min))"
    Write-Host "  -year                      Year"
    Write-Host "  -months                    Months (comma-separated of month numbers, e.g. 01,02,03 for Jan, Feb, Mar)"
}

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

