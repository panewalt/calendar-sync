calendar-sync
Script to copy calendar events between multiple calendars

Google calendars:
Using quickstart API code & instructions for Python, found here:
https://developers.google.com/google-apps/calendar/quickstart/python
Uses OAuth authentication.

Outlook calendars:
uses REST API, docs found here: https://stackoverflow.com/questions/31955710/creating-multiple-events-in-office365-rest-api
and here: https://msdn.microsoft.com/en-us/office/office365/api/calendar-rest-operations

The Problem: maintaining multiple calendars, for personal and work.
Google lets you "visually merge" calendars, i.e. you can see them on one screen.
This is not good enough.
Adding Outlook into the mix further complicates things.
When an event is placed on one calendar, I need it to show up on other calendars as scheduled time.
I don't want all calendars to see all details though - when an event is placed on one work calendar,
I need it to show up as "Busy" on all the other work calendars, and I want to see all details on
my personal calendar.

This script does that.  All calendars, Google and Outlook, are checked for 30 days into the future.
Any events found in a particular timeslot for a particular calendar are created as placeholders
in the other calendars.
Any event that is deleted from the original calendar on which it was placed will cause all the
placeholder events to be deleted also.
