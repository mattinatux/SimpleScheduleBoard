function Get-CountOfEvents {
    <# Counts remaining events for the week, Sunday through Saturday
    Outputs # of events for the current week #>
    [CmdletBinding()]
    param(
    [Parameter (Mandatory = $true)] $revisedEventList
    )

    $myObj = "" | Select-Object thisWeek,nextWeek

    $today = (Get-Date).Date
    # week start
    $monday = $today.AddDays(1 - $today.DayOfWeek.value__)
    # week end
    $sunday = $monday.AddDays(6)
    # next monday
    $nextMonday = $monday.AddDays(7)
    # next week end
    $nextSunday = $monday.AddDays(13)

    $i = 0
    $f = 0

    ForEach ($one in ($revisedEventList.datetime)) {
        $date = [datetime]$one
        if (($date -ge $monday) -and ($date -le $sunday)) {
            $i++
        } elseif (($date -ge $nextMonday) -and ($date -le $nextSunday)) {
            $f++
            write-output "alpha"
        }
    }

    $myObj.thisWeek = $i
    $myObj.nextWeek = $f

    $myObj
}
$updateTimeStamp = get-date -format HH:mm

$eventList = icalBuddy -ic 'Ptown Schedule' -f -nc -npn -nrd -nnc 20 -n -iep 'title,datetime,notes' -ps '|,|' -po 'datetime,title,notes' -b ',' -df '%b %e %Y' -eed eventsToday+120

# add column headers
$eventList = $eventList | ConvertFrom-Csv -Header e,datetime,title,notes | Select-Object datetime,title,notes

$excludes = @()
$excludes += 'Freestyle changed'
$excludes += 'Review and update Helen Health Fact'

$revisedEventList = @()
forEach ($eventy in $eventList) {
    if (!($eventy | Select-String -Pattern $excludes)) {
        $revisedEventList += $eventy
    }
    else {
        write-debug "$event matches exclusion filter"
    }
}

# Split datetime into three columns
$revisedEventList | ForEach-Object {
    $date = $_.datetime.split(" at ")[0]; 
    $time = $_.datetime.split(" at ")[1]; 
    $year = $date.substring($date.length - 4, 4);

    <# Remove the year from the date column, since it is going into it's own now.
    This regex just replaces the last 5 characters with nothing #>
    $date = $date -replace ".{5}$"

    <# Changes datetime column to a proper date but without time #>
    $_.datetime = ($date.Substring(5)) + " " + $year

    <# Substring 5 to remove blank space before the date. Trim wouldn't work. #>
    $_ | Add-Member -MemberType NoteProperty -Name date -Value ($date.Substring(5)); 
    $_ | Add-Member -MemberType NoteProperty -Name time -Value $time;
    $_ | Add-Member -MemberType NoteProperty -Name year -Value $year;
}

# Add DOW column
$revisedEventList | ForEach-Object {
    $dow = ((([datetime]$_.datetime).DayOfWeek).ToString().Substring(0,3)); 
    $_ | Add-Member -MemberType NoteProperty -Name dow -Value $dow
}

$orderedEventList = $revisedEventList | Select-Object dow,date,time,title,notes -First 13

<# Make the Schedule (event list) HTML Fragment to insert into the webpage #>

$orderedEventListHTML = $orderedEventList | ConvertTo-Html -Fragment -Property dow, date, time, title, notes


<# Working on Count of Events for this week and next week #>
$eventCounts = Get-CountOfEvents -revisedEventList $revisedEventList
$numThisWk = $eventCounts.thisWeek
$numNextWk = $eventCounts.nextWeek

$css = '
    .container {  display: grid;
    grid-template-columns: 180px 540px;
    grid-template-rows: 60px 180px 240px;
    gap: 1px 1px;
    grid-auto-flow: row;
    justify-content: center;
    align-content: stretch;
    justify-items: stretch;
    align-items: stretch;
    grid-template-areas:
        "numthiswk todaydate"
        "numthiswk schedule"
        "numnextwk schedule";
    }

    .schedule { grid-area: schedule; 
        background-color: #F8F8FF;
        padding: 5px;
    }

    table {
        border-spacing: 0px;
        font-weight: bold;
        width: 530px;
    }

    table td {padding: 7px;}

    /* tr:nth-child(even) { background-color:#FFEBCD; color: #00008B; }
    tr:nth-child(odd) { background-color:#F0FFF0; color: #2E8B57; } */

    tr:nth-child(odd) { background-color:black; color: yellow; font-weight: normal; }
    tr:nth-child(even) { background-color:white; color: black; }

    td:first-child {
        width: 48px;
    }
    td:nth-child(2) {
        width: 54px;
    }
    td:nth-child(3) {
        width: 72px;
    }

    .todaydate { grid-area: todaydate;
        background-color: #FFF;
        font-weight: bold;
        color: black;
        display: flex;
        justify-content: center;
        align-items: center;
        font-size: 1.8em;
        border-bottom: black;
        border-bottom-width: medium;
        border-bottom-style: solid;
    }

    .numthiswk { grid-area: numthiswk; 
        background-color: #FFB831;
        display: flex;
        justify-content: center;
        align-items: center;
        flex-wrap: wrap;
        padding: 16px;
    }

    .numnextwk { grid-area: numnextwk; 
        background-color: #00F3FF;
        display: flex;
        justify-content: center;
        align-items: center;
        flex-wrap: wrap;
        padding: 16px;
    }

    .numwkA {
        font-size: 1.25em;
        text-align: center;
        line-height: 1.2;
        font-weight: bold;
    }

    .numwkB {
        font-size: 7em;
        text-align: center;
    }
    '

$html = ""
$html += "<html><head><title>"
$html += "411 Board ($updateTimeStamp)"
$html += "</title>"
$html += "<style>$CSS</style>"

<# Old style didn't work on Safari / iPad OS #>
# $html += '<meta http-equiv="Refresh" content="3600">'

<# Sets refresh to every 2 minutes (1000 = 1s) #>
$html += "<script>function autoRefresh(){window.location=window.location.href;}setInterval('autoRefresh()',120000);</script>"

<# Moment, used for date stamp on page #>
$html += '<script src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js"></script>'
$html += "</head>"
$html += '<body><div class="container">
    <div class="schedule">'

<# SECTION A: SCHEDULE #>
<# The  REPLACE gets rid of ANSI color codes.
More here: https://old.reddit.com/r/PowerShell/comments/8z62js/powershell_core_on_linux_remove_color_codes_from/ #>
$html += $orderedEventListHTML -replace '\x1b\[[0-9;]*[a-z]', '' 

<# SECTION B: TODAY'S DATE #>
$html += '</div>
    <div class="todaydate">'

<# Removed in preference of using javascript and moment.js #>
# $html += "$todayFormatted"

$html += '<span id="date-time"></span>'
$html += '</div>
    <div class="numthiswk">'

<# SECTION C: NUMBER OF APPOINTMENTS THIS WEEK #>
$html += '<span class="numwkA">Appointments This Week</span>'
$html += '<span class="numwkB">'
$html += "$numThisWk</span>"
$html += '</div>
    <div class="numnextwk">'

<# SECTION D: NUMBER OF APPOINTMENTS NEXT WEEK #>
$html += '<span class="numwkA">Appointments <I>Next</I> Week</span>'
$html += '<span class="numwkB">'
$html += "$numNextWk</span>"
$html += '</div>
    </div>'

<# The following script block is for javascript on the page #>
$html += '<script>document.getElementsByTagName("tr")[0].remove();'
$html += "var dt = moment().format('dddd, MMMM Do');document.getElementById('date-time').innerHTML=dt;"
$html += '</script>'
    
$html += "</body></html>"
$html | Out-File $PSScriptRoot/index.html