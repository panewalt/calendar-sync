from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime, timedelta
import dateutil.parser

from myevent import MyEvent
from outlook import Outlook
    
class GCal:
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
    
    def __init__(self, ID, scope=None, appName=None, secretsFile=None, credentialsFile=None):
        self.ID = ID
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
        print('Getting the upcoming events for calendar %s' % self.ID)
        eventResult = self.service.events().list(
            calendarId='primary', timeMin=now, timeMax=then, singleEvents=True,
            orderBy='startTime').execute()
        eventList = eventResult.get('items', [])
        # now eventsList has a list of Google Calendar events.  Translate that to MyEvents events.
        myEventList = []
        for item in eventList:
            print("Google Event %s, start %s, end %s" % (item['summary'], item['start'], item['end']))
            if 'dateTime' not in item['start']: continue    # skip all-day events
            event = MyEvent()
            event.ID = item['id']
            event.summary = item['summary']
            event.start = item['start']['dateTime']
            event.end = item['end']['dateTime']
            if 'location' in item: event.location = item['location']
            if 'description' in item: event.description = item['description']
            #print("Event %s, start %s, end %s" % (event.summary, event.start, event.end))
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
        event = self.service.events().insert(calendarId='primary', body=GCalEvent).execute()
        return event

        
    def deleteEventFromCalendar(self, event):
        eventID = event.ID
        self.service.events().delete(calendarId='primary', eventId=eventID).execute()
        #print("Calendar %s: Deleted Event %s" % (self.ID, eventID))


        
def addEventsToMaster(eventList, ID, masterEventList):
    # master event list is a dictionary, with the start/end time as the key.
    # The value is another dictionary, with keys corresponding to the calendars.
    # That dictionary should ultimately have one entry for each calendar.
    for event in eventList:
        start = event.start  #.get('dateTime', event['start'].get('date'))
        end = event.end      #.get('dateTime', event['end'].get('date'))
        #print(start, end, event['summary'])
        eventKey = "%s-%s" % (start, end)       # create a key from start & end time
        if not eventKey in masterEventList:     # add it to the list, if necessary
            masterEventList[eventKey] = {}
        masterEventList[eventKey][ID] = event     # save this event in the timeslot
    return masterEventList


def findCalendarTag(event, calendars):
    for ID in calendars:
        tag = "<%s> " % ID
        if event.summary.startswith(tag):
            return ID
    return None

            
def main():
        
    calendars = {
        'PA':  {'type': 'Google', 'appName':'Calendar Sync', 'secrets':'pete_client_secret.json', 'creds_file':'pete-credentials.json', 'publishDetails': ''},
        'MOV': {'type': 'Google', 'appName':'MOV Calendar Sync', 'secrets':'mov_client_secret.json', 'creds_file':'mov-credentials.json', 'publishDetails': 'PA'},
        'GC':  {'type': 'Google', 'appName':'Calendar Sync', 'secrets':'gc_client_secret.json', 'creds_file':'gc-credentials.json', 'publishDetails': 'PA,MOV'},
        'UL':  {'type': 'Outlook', 'appName':'Calendar Sync', 'creds_file': 'ul-credentials.txt', 'publishDetails': 'PA'}
        }
    totalCalendars = len(calendars)
    masterEventList = {}

    for ID in calendars:
        cal = calendars[ID]
        if cal['type'] == 'Google':
            cal['instance'] = GCal(ID=ID, appName=cal['appName'], secretsFile=cal['secrets'], credentialsFile=cal['creds_file'])
        elif cal['type'] == "Outlook":
            cal['instance'] = Outlook(ID=ID, credentialsFile=cal['creds_file'])
        cal['eventList'] = cal['instance'].getEventsFromCalendar(daysAhead=30)
        masterEventList = addEventsToMaster(cal['eventList'], ID, masterEventList)
    
    for timeslot in masterEventList:
        if not 'T' in timeslot: continue    # skip all-day events

        print("======== Checking Timeslot %s:" % timeslot)
        # unpack the master events list - entries correspond to start/end times
        timeslotEvents = masterEventList[timeslot]     # it's a dictionary
        print("Timeslot has %s events" % len(timeslotEvents))
        # look at all the events in this timeslot, identify which calendars need placeholders added
        placeholderSet = set()    # use a set to hold the IDs of calendars that need placeholders
        primaryEventCalendarID = None
        for ID in calendars:
            if not ID in timeslotEvents:
                print("%s missing event for Calendar %s" % (timeslot, ID))
                placeholderSet.add(ID)
            else:
                event = timeslotEvents[ID]
                print(timeslot, ID, event.summary)
                if not primaryEventCalendarID and not findCalendarTag(event, calendars):
                    #print("Identified Primary Calendar %s" % ID)
                    primaryEventCalendarID = ID

        if primaryEventCalendarID:
            #for item in masterEventList[timeslot]:
            #    print(item, masterEventList[timeslot][item])
            # primary calendar for event identified - add events to others
            print("Primary Calendar for event: %s" % primaryEventCalendarID)
            publishDetails = calendars[primaryEventCalendarID]['publishDetails']
            event = timeslotEvents[primaryEventCalendarID]
            start = event.start
            end = event.end
            for ID in calendars:
                cal = calendars[ID]
                if ID in placeholderSet:     # this calendar needs a placeholder
                    if ID in publishDetails:
                        #cal['instance'].copyEvent(primaryEventCalendarID, event)
                        newEvent = MyEvent()
                        newEvent.createCopyOfEvent(primaryEventCalendarID, event)
                        print("Copying event %s from calendar %s to %s" % (event.summary, primaryEventCalendarID, ID))
                        input("Press Enter to continue")
                        result = cal['instance'].addEventToCalendar(newEvent)
                        print('Created Copied Event %s on Calendar %s' % (newEvent.summary, ID))
                        #exit(0)
                    else:
                        #cal['instance'].addPlaceholder(primaryEventCalendarID, start, end)
                        newEvent = MyEvent()
                        newEvent.createPlaceholderEvent(primaryEventCalendarID, start, end)
                        print("Adding placeholder %s to calendar %s" % (newEvent.summary, ID))
                        input("Press Enter to continue")
                        result = cal['instance'].addEventToCalendar(newEvent)
                        print('Created Placeholder Event %s on Calendar %s' % (newEvent.summary, ID))
                        #exit(0)
                                    
        else:
            # no primary calendar - this means the event was deleted from the primary calendar,
            # and all the ones remaining in this timeslot are placeholders.  Delete them.
            print("No Primary Calendar for any event in this timeslot - deleting placeholders...")
            for ID in calendars:
                cal = calendars[ID]
                if ID in timeslotEvents:
                    event = timeslotEvents[ID]
                    cal['instance'].deleteEventFromCalendar(event)
                    print("Deleted event %s from calendar %s" % (event.summary, ID))
                    input("Press Enter to continue")
            

if __name__ == '__main__':
    main()
