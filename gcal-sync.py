#from __future__ import print_function
#import httplib2
# Requires Python 3
import sys
if sys.version_info[0] < 3:
    print("Please start with python3")
    exit(1)
    
import os
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime, timedelta
import dateutil.parser

from myevent import MyEvent
from outlook import OutlookCalendar

gDaysAhead = 30
    
class GoogleCalendar:
    # class containing a Google calendar.
    # There will be several of these - we will consolidate events between them.
    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/CREDENTIALS_FILE
    scope = 'https://www.googleapis.com/auth/calendar'      #.readonly
    # secretsFile as entered when following this procedure:
    # https://developers.google.com/google-apps/calendar/quickstart/python
    # and modified with the name of a specific calendar, so we can have multiples
    # appName as entered in the Developer Console
    # https://console.developers.google.com
    # credentialsFile for this calendar, stored in ~/.credentials
    
    def __init__(self, calID, scope=None, appName=None, secretsFile=None, credentialsFile=None):
        self.calID = calID
        if not scope: scope = 'https://www.googleapis.com/auth/calendar'      #.readonly
        if not appName: appName = "Calendar Sync"
        self.credentials = self.getCredentials(scope, appName, secretsFile, credentialsFile)
        http = self.credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)
        # events_list = self.get_events_list()
        
        
    def getCredentials(self, scope, appName, secretsFile, credentialsFile):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        homeDir = os.path.expanduser('~')
        credentialsDir = os.path.join(homeDir, '.credentials')
        if not os.path.exists(credentialsDir):
            os.makedirs(credentialsDir)
        credentialsPath = os.path.join(credentialsDir, credentialsFile)

        store = Storage(credentialsPath)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(secretsFile, scope)
            flow.user_agent = appName
            credentials = tools.run(flow, store)
            print('Storing credentials to ' + credentialsPath)
        return credentials


    def getEventsFromCalendar(self, daysAhead=30):
        now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        then = (datetime.utcnow() + timedelta(days=daysAhead)).isoformat() + 'Z'
        #print('Getting the upcoming events for calendar %s' % self.calID)
        eventResult = self.service.events().list(
            calendarId='primary', timeMin=now, timeMax=then, singleEvents=True,
            orderBy='startTime').execute()
        eventList = eventResult.get('items', [])
        # now eventsList has a list of Google Calendar events.  Translate that to MyEvents events.
        myEventList = []
        for item in eventList:
            #print("%s: Google Event %s, start %s, end %s" % (self.calID, item['summary'], item['start'], item['end']))
            #if 'attendees' in item:
            #    print("Attendees: %s" % item['attendees'])
            if 'dateTime' not in item['start']: continue    # skip all-day events
            #print("%s: %s" % (self.calID, item))
            event = MyEvent()
            event.calID = self.calID
            event.ID = item['id']
            event.summary = item['summary']
            event.start = item['start']['dateTime']
            event.end = item['end']['dateTime']
            event.lastModified = item['updated']
            if 'location' in item: event.location = item['location']
            if 'description' in item: event.description = item['description']
            #print("Calendar: %s, Event %s, start %s, end %s, Location: %s, modified: %s" % (self.calID, event.summary, event.start, event.end, event.location, event.lastModified))
            myEventList.append(event)
        return myEventList


    def addEventToCalendar(self, event):
        # GCal events have Python datetimes for start & end times - create those
        GCalEvent = {}
        GCalEvent['summary'] = event.summary
        GCalEvent['start'] = {}
        GCalEvent['start']['dateTime'] = event.start    #datetime.strptime(event.start, "%Y-%m-%dT%H:%M:%S")
        GCalEvent['end'] = {}
        GCalEvent['end']['dateTime'] = event.end    #datetime.strptime(event.end, "%Y-%m-%dT%H:%M:%S")
        GCalEvent['location'] = event.location
        GCalEvent['description'] = event.description
        event = self.service.events().insert(calendarId='primary', body=GCalEvent).execute()
        return event

        
    def deleteEventFromCalendar(self, event):
        eventID = event.ID
        self.service.events().delete(calendarId='primary', eventId=eventID).execute()
        #print("Calendar %s: Deleted Event %s" % (self.calID, eventID))


        
def addEventsToMaster(eventList, calID, masterEventList):
    # master event list is a dictionary, with the start/end time as the key.
    # The value is a list, with an entry for each event in that timeslot.
    # That list should ultimately have at least one entry for each calendar,
    # and may have multiple entries for a given calendar.
    for event in eventList:
        start = event.start  #.get('dateTime', event['start'].get('date'))
        end = event.end      #.get('dateTime', event['end'].get('date'))
        #print(start, end, event['summary'])
        timeslot = "%s-%s" % (start, end)       # create a timeslot key from start & end time
        if not timeslot in masterEventList:             # if the timeslot doesn't yet exist,
            masterEventList[timeslot] = []              # create it
        masterEventList[timeslot].append(event)         # and save this event in the timeslot
            
    return masterEventList


def findCalendarTag(event):
    for calID in calendars:
        tag = "<%s> " % calID
        if event.summary.startswith(tag):
            #print("findCalendarTag: found tag for calendar %s in event %s" % (calID, event.summary))
            return calID
    return None


def getCalendarEvent(eventList, calID, summary=None):
    # given a list of events in a particular timeslot, look for an event on the specified calendar
    for event in eventList:
        if event.calID == calID:
            if not summary:
                return event
            elif event.summary == summary:
                return event
    return None


def getPrimaryEvent(eventList, calID):
    for event in eventList:
        if event.primary and event.calID == calID:
            return event
    return None
    
            
def main():
    global calendars
    calendars = {
        'PA':  {'type': 'Google', 'appName':'Calendar Sync', 'secrets':'pete_client_secret.json', 'creds_file':'pete-credentials.json', 'publishDetails': ''},
        'MOV': {'type': 'Google', 'appName':'MOV Calendar Sync', 'secrets':'mov_client_secret.json', 'creds_file':'mov-credentials.json', 'publishDetails': 'PA'},
        'GC':  {'type': 'Google', 'appName':'Calendar Sync', 'secrets':'gc_client_secret.json', 'creds_file':'gc-credentials.json', 'publishDetails': 'PA,MOV'},
        'UL':  {'type': 'Outlook', 'appName':'Calendar Sync', 'creds_file': 'ul-credentials.txt', 'publishDetails': 'PA'}
        }
    totalCalendars = len(calendars)
    masterEventList = {}

    print("Running at %s" % datetime.now())
    for calID in calendars:
        cal = calendars[calID]
        if cal['type'] == 'Google':
            cal['instance'] = GoogleCalendar(calID=calID, appName=cal['appName'], secretsFile=cal['secrets'], credentialsFile=cal['creds_file'])
        elif cal['type'] == "Outlook":
            cal['instance'] = OutlookCalendar(calID=calID, credentialsFile=cal['creds_file'])
        cal['eventList'] = cal['instance'].getEventsFromCalendar(daysAhead=gDaysAhead)
        eventsRetrieved = len(cal['eventList'])
        print("Calendar %s: Retrieved %d entries" % (calID, eventsRetrieved))
        if eventsRetrieved == 0:
            print("Exiting on 0 events retrieved from Calendar %s" % calID)
        masterEventList = addEventsToMaster(cal['eventList'], calID, masterEventList)

    for timeslot in masterEventList:
        if not 'T' in timeslot: continue    # skip all-day events
        # unpack the master events list - entries correspond to start/end times
        timeslotEvents = masterEventList[timeslot]     # it's a list
        #print("======== Checking Timeslot %s - total events: %d:" % (timeslot, len(timeslotEvents)))

        # look at all the events in this timeslot, identify which calendars need placeholders added
        primaryEventSet = set()
        for event in timeslotEvents:
            event.primary = not findCalendarTag(event)
            if event.primary and not event.summary.startswith("Canceled event:"):
                #print("Identified Primary Calendar %s for Event %s" % (event.calID, event.summary))
                primaryEventSet.add(event)

        if len(primaryEventSet) == 0:
            # no primary events - this means an event was deleted from the primary calendar,
            # and all the ones remaining in this timeslot are placeholders.  Delete them.
            print("No Primary Event in timeslot %s - deleting placeholders..." % timeslot)
            for event in timeslotEvents:
                calID = event.calID
                cal = calendars[calID]
                cal['instance'].deleteEventFromCalendar(event)
                print("Deleted event %s from timeslot %s, calendar %s" % (event.summary, timeslot, calID))
                #input("Press Enter to continue")
                continue

        # now go through all the primary events and add placeholders where needed.
        for primaryEvent in primaryEventSet:
            # primary calendar for event identified - add events to others
            publishDetails = calendars[primaryEvent.calID]['publishDetails']
            start = primaryEvent.start
            end = primaryEvent.end
            #print("Handling Primary Event %s on Calendar %s" % (primaryEvent.summary, primaryEvent.calID))
            for calID in calendars:
                cal = calendars[calID]
                if calID == primaryEvent.calID: continue    # primary event calendar doesn't need placeholders
                placeholderEvent = getCalendarEvent(timeslotEvents, calID)
                if placeholderEvent:
                    #print("Existing Placeholder Event %s on calendar %s" % (placeholderEvent.summary, calID))
                    if placeholderEvent.lastModified >= primaryEvent.lastModified:
                        continue
                    print("Placeholder Event %s on calendar %s is outdated, replacing it" % (placeholderEvent.summary, calID))
                    cal['instance'].deleteEventFromCalendar(placeholderEvent)
                    # and fall through to create the placeholder event
                # at this point, this calendar needs a placeholder
                newEvent = MyEvent()
                if calID in publishDetails:
                    newEvent.createCopyOfEvent(primaryEvent.calID, primaryEvent)
                else:
                    newEvent.createPlaceholderEvent(primaryEvent.calID, start, end)
                print("Adding event %s to calendar %s, timeslot %s" % (newEvent.summary, calID, timeslot))
                result = cal['instance'].addEventToCalendar(newEvent)
            

if __name__ == '__main__':
    main()
