# OutlookAddShabbatTimes
Add shabbat times to as Out of office meetings in Outlook calendar

```
  Usage: AddShabbatTimesToCalendar.ps1 [-city <city>] [-minutesRequiredBeforeShabbat <minutes>] [-year <year>] [-months <months>]
  
  Parameters:
    -city                      City code as defined at https://github.com/hebcal/dotcom/blob/master/hebcal.com/dist/cities2.txt
    -minutesRequiredBeforeShabbat   Number of minutes required before shabbat (travel home time + preparation) (default 60 min))
    -year                      Year
    -months                    Months (comma-separated of month numbers, e.g. 01,02,03 for Jan, Feb, Mar)
```

OutlookTools powershell module is taken from https://github.com/AmanDhally/OutlookTools