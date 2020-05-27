# public-misc

### OutlookBirthdayCalendar.ps1

I found Outlook.com's automatic birthday calendar creation incredibly unreliable, particularly when changing birthdays on a contact in iOS (Outlook account is added to iOS to allow contact syncing both directions, and Save Contacts is turned off in the Outlook iOS app)

Script:

1. Gets all contacts from desktop Outlook app (so you will need to setup your account there)
2. Set their birthdays to the vCard "no birthday" date of `1/01/4501 12:00:00 AM`
3. Do a send and receive and then wait for a bit for Outlook.com's backend to catch up
4. Set their birthdays back to the original values (which should prompt the backend to re-create the birthday calendar entry)
5. Delete the additional automatic recurring event that desktop Outlook creates in the normal calendar (titled `$subject's Birthday`)

Step 5 can be toggled off by setting `$deleteRecurringCalendarEntry = $false`
