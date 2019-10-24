
from dateutil import parser, tz


class MyEvent:

    def __init__(self, **kwargs):
        return self.createEvent(**kwargs)
        
    def createEvent(self, **kwargs):
        self.ID=kwargs.get('ID', '')
        self.summary=kwargs.get('summary', '')
        self.location=kwargs.get('location', '')
        self.description=kwargs.get('description', '')
        self.start=kwargs.get('start', '')
        self.end=kwargs.get('end', '')
        self.hangoutLink=kwargs.get('hangoutLink', '')
        self.conferenceData=kwargs.get('conferenceData', '')
        #self.blocked=kwargs.get('blocked', '')
        # now make sure dates are normalized.
        # Google dates come in as timezone-aware local datetime objects.
        # Outlook dates come in as UTC dates.
        # Make them both local objects
     

    def createPlaceholderEvent(self, calID, start, end):
        newSummary = "<%s> Busy" % calID
        self.createEvent(ID='', summary=newSummary, start=start, end=end)

    def createCopyOfEvent(self, calID, event):
        newSummary = "<%s> %s" % (calID, event.summary)
        #print("Creating copy of event - conferenceData: %s" % event.conferenceData)
        if hasattr(event, 'hangoutLink'): print("hangoutLink: %s" % event.hangoutLink)
        self.createEvent(ID=event.ID, summary=newSummary, location=event.location, description=event.description, start=event.start, end=event.end, hangoutLink=event.hangoutLink, conferenceData=event.conferenceData)
        #print("Event Copied - Summary: %s, Start: %s, End: %s, Location: %s" % (self.summary, self.start, self.end, self.location))
        #return newEvent
        
    def convertUTCtoLocalDatetime(self, s):
        # convert a time string in UTC format: 2017-08-02T14:00:00Z
        # to locally-aware datetime string:    2017-08-02T14:00:00+00:00
        dt = parser.parse(s)
        tzLocal = tz.tzlocal()
        local = dt.astimezone(tzLocal)
        dString = local.isoformat()
        return dString

    def convertToUTC(self, s):
        # convert a locally-aware datetime string:    2017-08-02T14:00:00+00:00
        # to a time string in UTC format:             2017-08-02T14:00:00Z
        dt = parser.parse(s)
        tzUTC = tz.tzutc()
        utc = dt.astimezone(tzUTC)
        dString = utc.isoformat()   
        #print("Input String: %s, Output: %s" % (s, dString))
        return dString
        
    
