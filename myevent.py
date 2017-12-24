
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
        # now make sure dates are normalized.
        # Google dates come in as timezone-aware local datetime objects.
        # Outlook dates come in as UTC dates.
        # Make them both local objects
     

    def createPlaceholderEvent(self, calID, start, end):
        newSummary = "<%s> Busy" % calID
        self.createEvent(ID='', summary=newSummary, start=start, end=end)

    def createCopyOfEvent(self, calID, event):
        newSummary = "<%s> %s" % (calID, event.summary)
        self.createEvent(ID=event.ID, summary=newSummary, location=event.location, description=event.description, start=event.start, end=event.end)
        #print("Event Copied - Summary: %s, Start: %s, End: %s, Location: %s" % (self.summary, self.start, self.end, self.location))
        #return newEvent
        
    def convertUTCtoLocalDatetime(self, s):
        # convert a time string in UTC format: 2017-08-02T14:00:00Z
        # to locally-aware datetime string:    2017-08-02T14:00:00+00:00
        utc = parser.parse(s)
        tzLocal = tz.tzlocal()
        local = utc.astimezone(tzLocal)
        dString = local.isoformat()
        return dString

    
