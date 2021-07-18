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


import googleapiclient.discovery
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class GoogleDrive(object):
    """Creates a channel to a google account's Drive documents.  The authenticated access is stored in self.service."""
    # Begin with no authenticated channel
    service = None
    # Scope of this Google Drive service
    SCOPE = 'https://www.googleapis.com/auth/drive.readonly'
    
    # MIME-types
    # Used to search for a specific type of file
    SHEETS = 'application/vnd.google-apps.spreadsheet'
    DOCS = None
    SLIDES = None
    PICTURES = None
    
    def __init__(self, clientSecretFile, credentialsFile, userAccount):
        # Setup the Drive API
        # To enable write-access if needed:
        #   go to Google Admin Page for omg.org (admin.google.com)
        #   Security -> Advanced Settings -> Manage API Client Access
        #   Enter the client id field from client_secret-omg-events.json for the TFCalendars project
        #   Enter 'https://www.googleapis.com/auth/drive' as the API scope, and Authorize
        SCOPES = 'https://www.googleapis.com/auth/drive.readonly'

        # Service Account Authorization, masquerades as omg-events@omg.org to get its calendars
        credentials = service_account.Credentials.from_service_account_file(clientSecretFile, scopes=[SCOPES], subject=userAccount)
        self.service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
        
    #
    # def getSheetsFiles(self):
    #     return self.getFilesOfType(GoogleDrive.SHEETS)

    def getFilesOfType(self, mimetype):
        """Return list of file descriptors"""
        logger.debug(self.service.files())
        results = self.service.files().list(q="mimeType='%s'" % (mimetype)).execute()
        items = results.get('files', [])
        logger.debug(items)
        if not items:
            logger.info('No files found.')
        else:
            logger.debug('Files:')
            for item in items:
                logger.debug(u'{0} ({1})'.format(item['name'], item['id']))
        return items
    
    def createFile(self, name="untitled", **kwargs):
        kwargs['name'] = name
        newFile = self.service.files().create(body=kwargs).execute()
        return newFile


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
    
    gs = GoogleDrive(os.path.abspath('client_secret-omg-events.json'), os.path.abspath('credentials.json'), "omg-events@omg.org")
    
    print(gs.getSheetsFiles())
