# Syncing Calendars

Some booking application exports their calendars in xlsx format (dont ask me why). We would like to import them into another application (nextcloud) as ics.

### First Step: Export 
Navigate to bookings -> list, export. 
The filter only works for the view. It always exports everything.

### 
Reduce to the rows of interest: resourceTitle="xxx", status Completed

reduce to columns of interest: bookingCode, Persons, price.groupinternal, status, bookingStartAtDate	bookingStartAtTime	bookingEndAtDate	bookingEndAtTime, customerNote
