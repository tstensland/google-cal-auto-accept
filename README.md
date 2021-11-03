# google-cal-auto-accept
Google Calendar Auto-Accept

Install and setup: 
1. Create an separate gmail account for the meeting rome. 
2. Open script.google.com
3. Select "New script"
4. Copy all *.html-files and code.gs into the script
5. Change calendarID and meetingRoomName
6. Name the project Autoreply-v1
7. Open triggers
8. Add new triggers:
    - choose which function to run
        - ProcessInvites
    - Deployment
        - Head
    - Event source
        - From calendar
    - Calendar details
        - Calendar updated
    - Calendar owner email
        - The email of the calendar (same as in point 5)
    - Failure notifications
        - Notify me immediately
    - Accept all required access to script/calendar


To use the scipt: 
1. Open Google calendar
2. Copy Secret adress in ical-format
3. Open Outlook and add the new calender under "Open calendar" - "From Internet"
