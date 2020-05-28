# how long to wait betweeen clearing and setting birthday
$waitTime = 60

# toggle this to $false to actually run
$whatIf = $true

# optional to turn off if you want to keep the additional recurring calendar entries in the normal calendar
$deleteRecurringCalendarEntry = $true

# default no birthday date of 4501
$noBirthdayDate = Get-Date('1/01/4501 12:00:00 AM')

$oldVerbosePreference = $VerbosePreference
$VerbosePreference = "continue"

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type] 
$outlook = New-Object -ComObject Outlook.Application -Verbose:$false
$namespace = $outlook.GetNamespace("MAPI")

# send and receive if outlook is in cached exchange mode
function Send-AndReceive {
    param(
        [switch]$force=$false
    )

    # if not running in whatif mode
    if (-not $whatIf -or $force) {
        # if cached exchange
        if ($outlook.session.Stores | ? { $_.IsCachedExchange }) {
            Write-Verbose "Send and receive"
            Start-Job { $namespace.SendAndReceive($false) } | Wait-Job -Timeout 15 | out-null
        }
    }
}

if ($whatIf) {
    Write-Warning "Running in whatIf mode - actions won't actually be performed"
}

Send-AndReceive -force

# get all contacts which have a birthday set which isn't the default no birthday date of 4501
$contacts = $outlook.session.GetDefaultFolder($olFolders::olFolderContacts).items | ? { $_.birthday -ne $noBirthdayDate -and $_.birthday } | sort Fullname
$contactsOriginal = ($contacts | select subject, birthday, @{N='SetBirthdayFrom';E={$_.birthday}}, @{N='SetBirthdayTo';E={$noBirthdayDate}})

Write-Verbose "Clearing birthdays"

# loop through contacts (twice, once to clear birthday, once to set it back again)
0..1 | % {
    $runCount = $_

    # if second run, then send and receive and then sleep to allow outlook.com's backend to catch up
    if ($runCount -eq 1) {
        Send-AndReceive
        Write-Verbose "Waiting for $waitTime seconds"
        if (-not $whatIf) {
            sleep $waitTime
        }
        Write-Verbose "Setting birthdays back to original values"
    }

    & {
        # loop through contacts
        foreach ($contact in $contacts) {
            $contactOriginal = $contactsOriginal | ? { $_.subject -eq $contact.subject }

            # second run - set back to original birthday
            if ($runCount -eq 1) {
                $contactOriginal.SetBirthdayFrom = $contactOriginal.SetBirthdayTo
                $contactOriginal.SetBirthdayTo = $contactOriginal.Birthday
            }

            # output what we are doing
            $contactOriginal | select Subject, SetBirthdayFrom, SetBirthdayTo
        
            # change birthday and save
            if (-not $whatIf) {
                $contact.Birthday = $contactOriginal.SetBirthdayTo
                $contact.save()
            }
        }
    } | ft -auto
}

Send-AndReceive

if ($deleteRecurringCalendarEntry) {
    Write-Verbose "Clearing recurring calendar entries"
}

# get recurring calendar entries
$cal = $outlook.session.GetDefaultFolder($olFolders::olFolderCalendar).items | ? { $_.IsRecurring }

# loop through contacts
& {
    foreach ($contact in $contacts) {
        # find calendar entries that match 
        if ($foundCalEntry = $cal | ? { $_.subject -eq "$($contact.subject)'s Birthday"} ) {
            $foundCalEntry | select subject, start

            # if calendar entries should be deleted
            if ($deleteRecurringCalendarEntry -and -not $whatIf) {
                $foundCalEntry.Delete()
            }
        }
    }
} | ft -auto

# dont need the final send and receive if cal entries weren't touched
if ($deleteRecurringCalendarEntry) {
    Send-AndReceive
}   

$VerbosePreference = $oldVerbosePreference