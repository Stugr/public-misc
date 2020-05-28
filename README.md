# public-misc

### OutlookBirthdayCalendar.ps1

Outlook.com has an automatic birthday calendar and I have found the creation of entries to be incredibly unreliable, particularly when changing birthdays on a contact in iOS (mail account is added to iOS to allow contact syncing in both directions, and Save Contacts is turned off in the Outlook iOS app)

Read about the birthday calendar: https://support.office.com/en-us/article/add-a-birthday-calendar-in-outlook-com-b8e636da-fda8-413f-940e-68396efa49a6

Script requires Outlook to be installed and the account you want to use to be setup. Untested with multiple accounts.

After running in default whatIf mode to ensure it's operating against the correct account and you're comfortable with what it will do, toggle `$whatIf = $true` to run it for real

1. Gets all contacts from desktop Outlook app (so you will need to setup your account there)
2. Set their birthdays to the vCard "no birthday" date of `1/01/4501 12:00:00 AM`
3. Do a send and receive and then wait for a bit for Outlook.com's backend to catch up
4. Set their birthdays back to the original values (which should prompt the backend to re-create the birthday calendar entry)
5. Delete the additional automatic recurring event that desktop Outlook creates in the normal calendar (titled `$subject's Birthday`)

Step 5 can be toggled off by setting `$deleteRecurringCalendarEntry = $false`
