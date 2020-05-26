$forceEditBirthday = $false
$deleteRecurringCalendarEntry = $false

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type] 
$outlook = New-Object -ComObject Outlook.Application

$contacts = $outlook.session.GetDefaultFolder($olFolders::olFolderContacts).items | ? { $_.birthday -ne '1/01/4501 12:00:00 AM'} | sort Fullname

# loop through contacts
foreach ($contact in $contacts) {
    $contact | select FullName, Birthday

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

$cal = $outlook.session.GetDefaultFolder($olFolders::olFolderCalendar).items | ? { $_.IsRecurring }

# find calendar entries that match 
foreach ($contact in $contacts) {
    if ($foundCalEntry = $cal | ? { $_.subject -eq "$($contact.fullname)'s Birthday"} ) {
        $foundCalEntry | select subject, start

        if ($deleteRecurringCalendarEntry) {

        }
    }
}
