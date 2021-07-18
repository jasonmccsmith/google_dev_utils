"""
Copyright Elemental Reasoning, LLC, 2019, 2020, 2021
All rights reserved unless otherwise specified in licensing agreement.
---------------
"""
import sys
import os
import os.path
import argparse

from errutils import erlogging
logger = erlogging.setup(lambda depth: sys._getframe(depth))
    

import datetime

import googleapiclient.discovery
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class GoogleCalendar(object):
    """Creates a channel to a google account's calendars.  The authenticated access is stored in self.service, and self.currentCal points to the actual calendar being accessed."""
    # Begin with no authenticated channel
    service = None
    # 'primary' is the name of the first 'standard' calendar created by a google account
    # If currentCal is never set to another value, this should always still work
    currentCal = 'primary'
    colors = None
    def __init__(self, clientSecretFile, credentialsFile, userAccount):
        # Setup the Calendar API
        # To enable write-access if needed:
        #   go to Google Admin Page for omg.org (admin.google.com)
        #   Security -> Advanced Settings -> Manage API Client Access
        #   Enter the client id field from client_secret-omg-events.json for the TFCalendars project
        #   Enter 'https://www.googleapis.com/auth/calendar' as the API scope, and Authorize
        SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'

        # Service Account Authorization, masquerades as omg-events@omg.org to get its calendars
        credentials = service_account.Credentials.from_service_account_file(clientSecretFile, scopes=[SCOPES], subject=userAccount)
        self.service = googleapiclient.discovery.build('calendar', 'v3', credentials=credentials)
        
        # # User Account Authorization.  Works to get all calendars, but *may* require a periodic authorization by logging into Google's auth request page with the omg-events@omg.org account.
        # creds = None
        # # The file token.pickle stores the user's access and refresh tokens, and is
        # # created automatically when the authorization flow completes for the first
        # # time.
        # if os.path.exists('token.pickle'):
        #     with open('token.pickle', 'rb') as token:
        #         creds = pickle.load(token)
        # # If there are no (valid) credentials available, let the user log in.
        # if not creds or not creds.valid:
        #     if creds and creds.expired and creds.refresh_token:
        #         creds.refresh(Request())
        #     else:
        #         flow = InstalledAppFlow.from_client_secrets_file('credentials.json', [SCOPES])
        #         creds = flow.run_local_server()
        #     # Save the credentials for the next run
        #     with open('token.pickle', 'wb') as token:
        #         pickle.dump(creds, token)
        # self.service = googleapiclient.discovery.build('calendar', 'v3', credentials=creds)

    def getCalendars(self, only=None):
        """Return list of calendar descriptors"""
        cals = self.service.calendarList().list().execute().get('items', [])
        if only:
            cals = [ c for c in cals if c['summary'] == only ]
        return cals
    
    def getCalendarObjectForCalendarNamed(self, calName):
        cals = [ c for c in self.service.calendarList().list().execute().get('items', []) if c['summary'].lower() == calName.lower() ]
        if len(cals) == 1:
            return cals[0]
        elif len(cals) > 1:
            logger.error("Something serious wrong with Google calendar, returned more than one calendar named %s" % (calName))
        else:
            logger.error("Calendar named '%s' was not found" % (calName))
        return None
        
    def getCalendarIDForCalendarNamed(self, calName):
        cal = self.getCalendarObjectForCalendarNamed(calName)
        if cal:
            return cal['id']
        return "NO-SUCH-CALENDAR"
        
    def getColorForCalendarNamed(self, calName):
        if self.colors == None:
            logger.info("Getting colors")
            self.colors = self.service.colors().get().execute()
        cal = self.getCalendarObjectForCalendarNamed(calName)
        colId = cal['colorId']
        # print(colId)
        # print(self.colors['calendar'])
        if colId in self.colors['calendar']:
            return self.colors['calendar'][colId]['background']
        return "NO-SUCH-CALENDAR"
        
    def getCanonicalCalendarName(self, calName):
        cal = self.getCalendarObjectForCalendarNamed(calName)
        if cal:
            return cal['summary']
        return "NO-SUCH-CALENDAR"
        
        
    def getNextNEvents(self, count):
        """Get the next 'count' events on the calendar, after current date and time"""
        # Call the Calendar API
        logger.info('Getting the upcoming %d events' % (count))
        now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        events_result = self.service.events().list(calendarId=self.currentCal, timeMin=now,
                                              maxResults=count, singleEvents=True,
                                              orderBy='startTime').execute()
        logger.debug("events_result: %r" % (events_result))
        events = events_result.get('items', [])

        if not events:
            logger.info('No upcoming events found.')
        else:
            logger.debug(events)
        return events
    
    def getEventsInDateTimeRange(self, dateStart, dateEnd):
        """Get all events starting at dateStart, through dateEnd.
        dateStart and dateEnd are datetime.datetime objects, whose time is respected.
        They must have a valid timezone set for the location where the event is happening.
        Timezone naive datetimes are not allowed."""
        # Call the Calendar API
        logger.debug('Getting events from calendar: %s' % (self.currentCal))
        logger.debug('Getting the events from %r to %r' % (dateStart, dateEnd))
        logger.debug('Getting the events from %r to %r' % (dateStart.isoformat(), dateEnd.isoformat()))
        events_result = self.service.events().list(calendarId=self.currentCal, timeMin=dateStart.isoformat(),
                                              timeMax=dateEnd.isoformat(), singleEvents=True,
                                              orderBy='startTime', maxResults=250).execute()
        logger.debug("events_result: %r" % (events_result))
        events = events_result.get('items', [])
        if not events:
            logger.info('No upcoming events found.')
            events = []
        else:
            logger.debug(events)
        if events_result.get('nextPageToken') != None:
            logger.info("THERE ARE MORE!")

        return events

    def getEventsInDateRange(self, start, end):
        """Get all events starting at 00:00:00 on start, through 23:59:59 on end.
        start and end are datetime.datetime objects, whose time is ignored.
        They must have a valid timezone set for the location where the event is happening,
        for the 0000-2359 range to make sense.  Timezone naive datetimes are not allowed."""
        # Call the Calendar API
        logger.debug('Getting the events from %r to %r' % (start, end))
        return self.getEventsInDateTimeRange(start.replace(hour=0, minute=0), end.replace(hour=23, minute=59))

class GoogleCalendarMutable(GoogleCalendar):
    def __init__(self, clientSecretFile, credentialsFile, userAccount):
        SCOPES = 'https://www.googleapis.com/auth/calendar'

        # Service Account Authorization, masquerades as omg-events@omg.org to get its calendars
        credentials = service_account.Credentials.from_service_account_file(clientSecretFile, scopes=[SCOPES], subject=userAccount)
        self.service = googleapiclient.discovery.build('calendar', 'v3', credentials=credentials)
        
    def deleteEvent(self, eventID):
        results = self.service.events().delete(calendarId=self.currentCal, eventId=eventID).execute()
        return results

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description = 'Google Calendar Authenticator')
    # Debugging and testing
    parser.add_argument('--verbose', default=False, action='store_true')
    parser.add_argument('--debug', default=False, action='store_true')
    args = parser.parse_args()

    # Logging to file is always set to debug, this is just for console
    if args.verbose:
        logger.setLevel(erlogging.INFO)
    if args.debug:
        logger.setLevel(erlogging.DEBUG)

