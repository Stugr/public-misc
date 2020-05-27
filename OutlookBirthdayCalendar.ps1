$forceEditBirthday = $false
$deleteRecurringCalendarEntry = $false

# default no birthday date of 4501
$noBirthdayDate = '1/01/4501 12:00:00 AM'

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type] 
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# send and receive if outlook is in cached exchange mode
function Send-AndReceive {
    if ($outlook.session.Stores | ? { $_.IsCachedExchange }) {
        Start-Job { $namespace.SendAndReceive($false) } | Wait-Job -Timeout 15 | out-null
    }
}

Send-AndReceive

# get all contacts which have a birthday set which isn't the default no birthday date of 4501
$contacts = $outlook.session.GetDefaultFolder($olFolders::olFolderContacts).items | ? { $_.birthday -ne $noBirthdayDate -and $_.birthday } | sort Fullname

# loop through contacts
& {
    foreach ($contact in $contacts) {
        $contact | select Subject, Birthday

        # edit birthday to force re-creation of birthday calendar entry
        if ($forceEditBirthday) {
            # add a day to their birthday
            $contact.Birthday = ($contact.Birthday).AddHours(24)
            $contact.save()

            # remove a day from their birthday (resetting it back to what it was originally)
            $contact.Birthday = ($contact.Birthday).AddHours(-24)
            $contact.save()
        }
    }
} | ft -auto

Send-AndReceive

# get recurring calendar entries
$cal = $outlook.session.GetDefaultFolder($olFolders::olFolderCalendar).items | ? { $_.IsRecurring }

# loop through contacts
& {
    foreach ($contact in $contacts) {
        # find calendar entries that match 
        if ($foundCalEntry = $cal | ? { $_.subject -eq "$($contact.subject)'s Birthday"} ) {
            $foundCalEntry | select subject, start

            # delete cal entry
            if ($deleteRecurringCalendarEntry) {
                $foundCalEntry.Delete()
            }
        }
    }
} | ft -auto

Send-AndReceive
