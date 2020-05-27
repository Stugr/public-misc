# how long to wait betweeen clearing and setting birthday
$waitTime = 60

# default no birthday date of 4501
$noBirthdayDate = Get-Date('1/01/4501 12:00:00 AM')

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type] 
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# send and receive if outlook is in cached exchange mode
function Send-AndReceive {
    if ($outlook.session.Stores | ? { $_.IsCachedExchange }) {
        ""
        "Send and receive"
        Start-Job { $namespace.SendAndReceive($false) } | Wait-Job -Timeout 15 | out-null
    }
}

Send-AndReceive

# get all contacts which have a birthday set which isn't the default no birthday date of 4501
$contacts = $outlook.session.GetDefaultFolder($olFolders::olFolderContacts).items | ? { $_.birthday -ne $noBirthdayDate -and $_.birthday } | sort Fullname
$contactsOriginal = ($contacts | select subject, birthday)

# loop through contacts (twice, once to clear birthday, once to set it back again)
"Clearing birthdays"
0..1 | % {
    $runCount = $_

    # if second run, then send and receive and then sleep to allow outlook.com's backend to catch up
    if ($runCount -eq 1) {
        Send-AndReceive
        "Waiting for $waitTime seconds"
        sleep $waitTime
        "Setting birthdays back to original values"
        ""
    }

    # loop through contacts
    foreach ($contact in $contacts) {
        # first run - clear birthday
        if ($runCount -eq 0) {
            $setBirthday = $noBirthdayDate
        }
        # second run - set back to original birthday
        else {
            $setBirthday = ($contactsOriginal | ? { $_.subject -eq $contact.subject }).birthday
        }

        # output what we are doing
        $contact | select Subject, Birthday, @{N='SetBirthdayTo';E={$setBirthday}}
            
        # change birthday and save
        $contact.Birthday = $setBirthday
        $contact.save()
    }
}

Send-AndReceive

# get recurring calendar entries
$cal = $outlook.session.GetDefaultFolder($olFolders::olFolderCalendar).items | ? { $_.IsRecurring }

# loop through contacts
& {
    foreach ($contact in $contacts) {
        # find calendar entries that match 
        if ($foundCalEntry = $cal | ? { $_.subject -eq "$($contact.subject)'s Birthday"} ) {
            $foundCalEntry | select subject, start

            # delete calendar entry
            $foundCalEntry.Delete()
        }
    }
} | ft -auto

Send-AndReceive
