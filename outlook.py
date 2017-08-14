import os
import requests
import json
import base64

from datetime import datetime, timedelta, timezone
from dateutil import parser, tz
from myevent import MyEvent

# Outlook Calendar access via REST API
# taken from info found here: https://stackoverflow.com/questions/31955710/creating-multiple-events-in-office365-rest-api
# and here: https://msdn.microsoft.com/en-us/office/office365/api/calendar-rest-operations
# Note: there is an OAuth procedure which I probably should have used, but it was easier to do it this way.
# Username and password are hardcoded in here and in gcal-sync.py.

# Outlook likes to return UTC date/times.  Google returns local date/times, which are easier to work with,
# so Outlook events get converted from UTC to local times.
# Once done, we can create new events in here using the local times.

class Outlook:
    def __init__(self, ID, credentialsFile=None):
        # Set the request parameters - this URL is the main URL for creating and deleting events
        # it also works for getting events, but I found that the calendarView function is better, because it allows dates to be specified.
        # so that one is defined below, in getEventsFromCalendar.
        self.baseUrl = 'https://outlook.office365.com/api/v1.0/me/events'
        self.ID = ID
        self.getCredentials(credentialsFile)

    def getCredentials(self, credentialsFile):
        homeDir = os.path.expanduser('~')
        credentialsDir = os.path.join(homeDir, '.credentials')
        if not os.path.exists(credentialsDir):
            os.makedirs(credentialsDir)
        credentialsPath = os.path.join(credentialsDir, credentialsFile)
        with open(credentialsPath, 'rt') as fh:
            self.user = fh.readline()
            self.pwd = fh.readline()
            #print(self.user, self.pwd)
            
    def createRequestHeaders(self):
        # the below mess properly encodes the authorization string in Python 3:
        upw = ('%s:%s' % (self.user, self.pwd)).replace('\n', '')
        auth = base64.b64encode(upw.encode()).decode('ascii')
        headers = {'Authorization': 'Basic %s' % auth, 'Content-Type': 'application/json', 'Accept': 'application/json'}
        return headers

    def getEventsFromCalendar(self, daysAhead=30):
        headers = self.createRequestHeaders()
        now = datetime.utcnow().replace(microsecond=0).isoformat() + 'Z' # 'Z' indicates UTC time
        then = (datetime.utcnow().replace(microsecond=0) + timedelta(days=daysAhead)).isoformat() + 'Z'
        #url = "https://outlook.office365.com/api/v1.0/me/events?$Select=Subject,Start,End,Location&start=%s&end=%s" % (now, then)
        url = "https://outlook.office365.com/api/v1.0/me/calendarView?$Select=Subject,Start,End,Location,BodyPreview&$top=250&startDateTime=%s&endDateTime=%s" % (now, then)
        #print(url)
        r = requests.get(url, headers=headers)
        returnDict = json.loads(r.text)
        eventList = returnDict['value']

        myEventList = []
        for item in eventList:
            if item['Start'] < now:
                #print("Skipping event in the past (%s, %s)" % (item['Subject'], item['Start']))
                continue
            #print("Outlook Item %s, start %s, end %s" % (item['Subject'], item['Start'], item['End']))
            #if 'dateTime' not in item['start']: continue    # skip all-day events
            event = MyEvent()
            event.ID = item['Id']
            event.summary = item['Subject']
            # Outlook dates are provided in UTC.  Convert to local-aware datetime.
            event.start = event.convertUTCtoLocalDatetime(item['Start'])
            event.end = event.convertUTCtoLocalDatetime(item['End'])
            event.location = item['Location']
            event.description = item['BodyPreview']
            #if 'description' in item: event.description = item['description']
            #print("Outlook Event %s, start %s, end %s" % (event.summary, event.start, event.end))
            myEventList.append(event)
        return myEventList


    def addEventToCalendar(self, event):
        # Create JSON payload
        '''
        # Test Outlook event:
        OutlookEvent = {
          "Subject": "My Subject",
          "Body": {
            "ContentType": "HTML",
            "Content": ""
          },
          "Start": "2015-08-11T07:00:00-05:00",
          "StartTimeZone": "Central Standard Time",
          "End": "2015-08-11T15:00:00-05:00",
          "EndTimeZone": "Central Standard Time",
        }
        '''        
        OutlookEvent = {
            'Subject': event.summary,
            #'Location': event.location,
            'Body':  {"ContentType": "HTML", "Content": event.description},
            'Start': event.start,
            #'StartTimeZone': 'MDT', #'Mountain Daylight Time',
            'End': event.end    #,
            #'EndTimeZone': 'MDT'    #'Mountain Daylight Time'
        }
        if event.location != '':
            OutlookEvent['Location'] = event.location

        json_payload = json.dumps(OutlookEvent)
        #print(json_payload)
        headers = self.createRequestHeaders()
        r = requests.post(self.baseUrl, headers=headers, data=json_payload)
        #print("Outlook Event created: %s" % event.summary)
        #print(r.text)

        
    def deleteEventFromCalendar(self, event):
        eventID = event.ID
        headers = self.createRequestHeaders()
        url = "%s/%s" % (self.baseUrl, eventID)
        #print(url)
        r = requests.delete(url, headers=headers)
        #print("Calendar %s: Deleted Event %s" % (self.ID, eventID))

        
def main():

    outlook = Outlook("ID", "ul-credentials.txt")
    print("Retrieving events from calendar")
    outlook.getEventsFromCalendar(daysAhead=30)

    event = MyEvent(ID="Test", summary="Test Event")
    startTime = datetime.utcnow().replace(microsecond=0)
    startStr = startTime.isoformat() + 'Z' # 'Z' indicates UTC time
    print(startStr)
    event.start = event.convertUTCtoLocalDatetime(startStr)
    print("Adding event to calendar - start time %s" % startStr)
    input("Press Enter to do it")
    endStr = (startTime + timedelta(hours=1)).isoformat() + 'Z'
    event.end = endStr
    outlook.addEventToCalendar(event)
    
if __name__ == '__main__':
    main()


