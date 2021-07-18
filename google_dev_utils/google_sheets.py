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

import unittest
from datetime import timedelta, datetime
import time

import googleapiclient.discovery
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from google_drive import GoogleDrive

class NoSheetFile(Exception):
    pass

def intToSheetsCol(colInt):
    """1 -> A, 26 -> Z, 27 -> AA, 52 -> AZ, etc"""
    if colInt < 1:
        logger.error("Bad value passed to intToSheetsCol: %d, must be greater than 0" % (colInt))
        return ""
    # def conv(num,b):
    #     convStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    #     if num<=b:
    #         return convStr[num-1]
    #     else:
    #         return conv(num//b,b) + convStr[(num%b)-1]
    # return conv(colInt, 26)
    
    divisor = int(colInt/26)
    remainder = (colInt % 26)
    if remainder == 0:
        finalChar = 'Z'
    else:
        finalChar = chr(remainder + 64)
    if remainder == 0:
        divisor -= 1
    if divisor == 0:
        return finalChar
    return chr(divisor + 64) + finalChar

def sheetsColToInt(colStr):
    """A -> 1, Z -> 26, AA -> 27, AZ -> 52, etc"""
    if colStr == "":
        logger.error("Bad value passed to sheetsColToInt: must not be empty string")
        return -1
    col = 0
    # Reverse string and ensure uppercase
    colStr = colStr[::-1].upper()
    for i, c in enumerate(colStr):
        col +=  (ord(c) - 64) * (26 ** i)
    return col

def splitCR(cr):
    """Returns 1-indexed numeric col/row equivalents of cell range format single cell position as tuple (col, row).
    Really not sure a regex would be faster or clearer..."""
    logger.debug(str(cr))
    if len(cr) == 0:
        return -1, -1
    col = ""
    row = ""
    for c in cr:
        if c.isalpha():
            col += c
        else:
            row += c
    logger.debug("col: {}, row: {}".format(col, row))
    if col == "" or row == "":
        logger.error("Malformed colRow: %r" % (cr))
    return sheetsColToInt(col), int(row)

def unSparse2D(data):
    """Given a jagged 2D data set, fill out with empty strings to be flush right"""
    maxLen = max([len(row) for row in data])
    for row in data:
        if len(row) < maxLen:
            for i in range(len(row), maxLen):
                row.append("")
    return data

def ensure2D(data):
    if type(data) is list:
        if len(data) > 0:
            if type(data[0]) is list:
                pass
            else:
                logger.debug("Wrapping list in list to satisfy update API.")
                data = [data]
        else:
            logger.warning("Data to writeData must be 2D list, sent empty data list, appending empty list.")
            data.append([])
    else:
        logger.warning("Data to writeData must be 2D list, wrapping raw data in 2D lists.")
        data = [[data]]
    return data

def trimList(l):
    if len(l) > 0:
        while l[-1] == "":
            l = l[:-1]
            if len(l) == 0:
                break
    return l



class GoogleSheets(object):
    """Creates a channel to a google account's sheets documents.  The authenticated access is stored in self.service, and self.currentDoc points to the actual sheets document being accessed."""
    # Begin with no authenticated channel
    service = None
    # 'primary' is the name of the first 'standard' calendar created by a google account
    # If currentCal is never set to another value, this should always still work
    currentDoc = 'primary'
    mimeType = 'application/vnd.google-apps.spreadsheet'
    def __init__(self, clientSecretFile, credentialsFile, userAccount):
        # Setup the Sheets API
        # To enable write-access if needed:
        #   go to Google Admin Page for omg.org (admin.google.com)
        #   Security -> Advanced Settings -> Manage API Client Access
        #   Enter the client id field from client_secret-omg-events.json for the TFCalendars project
        #   Enter 'https://www.googleapis.com/auth/spreadsheets' as the API scope, and Authorize
        # SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
        SCOPES = 'https://www.googleapis.com/auth/spreadsheets'

        # Service Account Authorization, masquerades as omg-events@omg.org to get its calendars
        credentials = service_account.Credentials.from_service_account_file(clientSecretFile, scopes=[SCOPES], subject=userAccount)
        self.service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)
        
        self.drive = GoogleDrive(clientSecretFile, credentialsFile, userAccount)
        
        self.sheet = None

    def getSheets(self):
        """Return list of sheets descriptors"""
        items = self.drive.getFilesOfType(GoogleSheets.mimeType)
        logger.debug("Sheets: %d" % (len(items)))
        return items
        
    def getSheet(self, name):
        """Return sheet named 'name'"""
        sheet = None
        for s in self.getSheets():
            logger.debug("Sheet: '%s'" % (s['name']))
            if s['name'] == name:
                sheet = s['id']
                break
        return sheet
        
    def createSheetFile(self, name):
        data = {'properties': {'title': name}}
        result = self.service.spreadsheets().create(body=data).execute()
        logger.debug(result)
        return result['spreadsheetId']
    
    def getOrCreateSheet(self, name, doNotWrite=False):
        """Return the *ID* of the sheet, not the sheet itself.  Send this ID into the Sheet constructor."""
        sheet = self.getSheet(name)
        if sheet == None and not doNotWrite:
            sheet = self.createSheetFile(name)
        logger.info(sheet)
        return sheet
    

class GoogleSheetAtomic(object):
    
    googleReadLimit = 100
    googleReadTimeframe = 100
    googleWriteLimit = 100
    googleWriteTimeframe = 100

    def __init__(self, sheetsAccessor, name="", sheetID="", createIfAbsent=False):
        logger.debug("Building from sheet: '{}'".format(name))
        # Name property data store
        self._name = None
        self.sheet = None
        self.colFirst = True
        self.sheets = sheetsAccessor
        if name != "" and sheetID != "":
            logger.error("Provide name OR ID of Google Sheet to retrieve, not both")
        if name != "":
            self.name = name
            self.sheet = self.sheets.getOrCreateSheet(name)
        elif sheetID != "":
            self.sheet = sheetID
        else:
            self.sheet = None  # Not yet associated with any sheet on google
            self.name = None
        self.data = None
        self.dataMaxLen = 0
        # TODO Cache both kinds, or does that lead to inconsistencies
        self.dataByRow = None
        self.dataByCol = None
        self.dataIsClean = False
        self.headerRow = False
        
        self.readCount = 0
        self.writeCount = 0
        self.readTimer = datetime.now() + timedelta(seconds=GoogleSheetAtomic.googleReadTimeframe)
        self.writeTimer = datetime.now() + timedelta(seconds=GoogleSheetAtomic.googleWriteTimeframe)
    
    @property
    def name(self):
        "The name property."
        if self._name == None and self.sheet != None:
            # fetch name from sheet online
            pass
        return self._name
    
    @name.setter
    def name(self, value):
        if self._name != value:
            # change name online
            pass
        self._name = value
        
    
    @name.deleter
    def name(self):
        del self._name
    
    
    @property
    def maxCols(self):
        "The maxCols property."
        return max( [ len(row) for row in self.data] )

    def setupDefault(self):
        pass

    def create(self):
        # Create new spreadsheet online
        # Check to make sure it doesn't exist
        for s in s.self.sheets.getSheets():
            if s['id'] == self.sheet:
                logger.error("Cannot create sheet %s, one with the same ID already exists." % (self.name))
    
    def makeBigBoldText(self, cellRange):
        """Given a 1-indexed cellRange, make it bold 12pt font"""
        # Decompose cellRange to col and row, then adjust to 0-indexed for API
        if ':' in cellRange:
            start, end = cellRange.split(':')
            startCol, startRow = splitCR(start)
            startCol -= 1
            startRow -= 1
            endCol, endRow = splitCR(end)
        else:
            startCol, startRow = splitCR(cellRange)
            endCol = startCol
            endRow = startRow
            startCol -= 1
            startRow -= 1
        logger.debug("Adjusting cell range %r - API as (%r, %r) -> (%r, %r) - to 12pt bold" % (cellRange, startRow, startCol, endRow-1, endCol-1))
        sheetId = 0
        requests = [ { \
              "repeatCell": { \
                "range": { \
                  "sheetId": sheetId, \
                  "startRowIndex": startRow, \
                  "endRowIndex": endRow, \
                  "startColumnIndex": startCol, \
                  "endColumnIndex": endCol \
                }, \
                "cell": { \
                  "userEnteredFormat": { \
                    "wrapStrategy": "OVERFLOW_CELL", \
                    "horizontalAlignment" : "LEFT", \
                    "textFormat": { \
                      "fontSize": 12, \
                      "bold": True \
                    } \
                  } \
                }, \
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,wrapStrategy)" \
              } \
            } ]
        
        logger.debug(requests)

        body = { 'requests': requests }
        result = self.batchUpdate(body)
        return result
        
    def makeHeaderRow(self):
        sheetId = 0
        requests = [ { \
              "repeatCell": { \
                "range": { \
                  "sheetId": sheetId, \
                  "startRowIndex": 0, \
                  "endRowIndex": 1 \
                }, \
                "cell": { \
                  "userEnteredFormat": { \
                    "backgroundColor": { \
                      "red": 0.5, \
                      "green": 0.5, \
                      "blue": 0.5 \
                    }, \
                    "horizontalAlignment" : "CENTER", \
                    "textFormat": { \
                      "foregroundColor": { \
                        "red": 0.0, \
                        "green": 0.0, \
                        "blue": 0.0 \
                      }, \
                      "fontSize": 12, \
                      "bold": True \
                    } \
                  } \
                }, \
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)" \
              } \
            }, \
            { \
              "updateSheetProperties": { \
                "properties": { \
                  "sheetId": sheetId, \
                  "gridProperties": { \
                    "frozenRowCount": 1 \
                  } \
                }, \
                "fields": "gridProperties.frozenRowCount" \
              } \
            } ]

        body = { 'requests': requests }
        logger.debug("Setting row 1 as header row")
        self.headerRow = True
        result = self.batchUpdate(body)
        return result
        
    def sortOnCol(self, col):
        """Column here is *ZERO* indexed... can send in index from colMap directly.
        Assume that row 1 is a frozen header and not to be sorted."""
        sheetId = 0
        requests = [ { "sortRange": \
                        { "range": { \
                              "sheetId": sheetId, \
                              "startRowIndex": 1, \
                              "startColumnIndex": 0, \
                            }, \
                            "sortSpecs": [ \
                              { \
                                "dimensionIndex": col, \
                                "sortOrder": "ASCENDING" \
                              } \
                            ] \
                          } \
                     } ]
        body = { 'requests': requests }
        logger.debug("Sorting on column %d (%s)" % (col, intToSheetsCol(col+1)))
        result = self.batchUpdate(body)
        self.sortOnColCache(col)
        return result
    
    def sortOnColCache(self, col):
        logger.debug("Sorting cache on column %d (%s)" % (col, intToSheetsCol(col+1)))
        logger.debug("Pre-sorted data...")
        logger.debug(self.data)
        if self.colFirst:
            logger.error("Is not implemented for colFirst mode")
            pass
        else:
            # logger.debug("{}".format(self.data))
            if self.headerRow:
                logger.debug("Locking header row")
                # logger.debug("Sorted remainder: {}".format(sorted(self.data[1:], key=lambda x: x[col])))
                self.data = [self.data[0]] + sorted(self.data[1:], key=lambda x: x[col])
            else:
                self.data.sort(key=lambda x: x[col])
        logger.debug("Sorted data...")
        logger.debug(self.data)
        
    
    def insertBlankRowAfter(self, row):
        """Column here is *ZERO* indexed... can send in index from colMap directly."""
        self.insertBlankRowsAfter(row, row+1)
    
    def insertBlankRowsAfter(self, row, end):
        """Column here is *ZERO* indexed... can send in index from colMap directly. End row is exclusive, one past intended last row added."""
        sheetId = 0
        requests = [ { \
                          "insertDimension": { \
                            "range": { \
                              "sheetId": sheetId, \
                              "dimension": "ROWS", \
                              "startIndex": row, \
                              "endIndex": end \
                            }, \
                            "inheritFromBefore": False \
                          } \
                     } ]
        body = { 'requests': requests }
        logger.debug("Inserting new blank row after row %d" % (row))
        result = self.batchUpdate(body)
        self.insertBlankRowsAfterCache(row, end)
        return result
    
    def insertBlankRowsAfterCache(self, row, end):
        # logger.debug("self.data: {}".format(self.data))
        if self.colFirst:
            logger.error("Not yet implemented for colFirst mode")
        else:
            for r in range(row, end):
                logger.debug("Inserting blank (empty) row after {}".format(r))
                # Adjust r because insert inserts *before* the index, not after
                self.data.insert(r, ["" for _ in range(max([len(r2) for r2 in self.data])) ] )
    
        
    def manageReadRate(self):
        self.readCount += 1
        # Give it a little wiggleroom
        if self.readCount >= GoogleSheetAtomic.googleReadLimit - 2:
            # wait until write timer elapses
            pauseSecs = (self.readTimer - datetime.now()).seconds
            if pauseSecs > 0:
                logger.warning("Pausing %d seconds to not run over google read rate limits..." % (pauseSecs))
                time.sleep(pauseSecs)
            self.readTimer = datetime.now() + timedelta(seconds=GoogleSheetAtomic.googleReadTimeframe)
            self.readCount = 1
    
    def manageWriteRate(self):
        self.writeCount += 1
        # Give it a little wiggleroom
        if self.writeCount >= GoogleSheetAtomic.googleWriteLimit - 2:
            # wait until write timer elapses
            pauseSecs = (self.readTimer - datetime.now()).seconds
            if pauseSecs > 0:
                logger.warning("Pausing %d seconds to not run over google write rate limits..." % (pauseSecs))
                time.sleep(pauseSecs)
            self.writeTimer = datetime.now() + timedelta(seconds=GoogleSheetAtomic.googleWriteTimeframe)
            self.writeCount = 1
    
    # def __updateData(self, cellRange, data, majDim="ROWS"):
        
    def writeData(self, cellRange, data, majDim="ROWS"):
        # check to see if data is 2D list
        data = ensure2D(data)
        logger.debug("Updating: '%s' with data: %r" % (cellRange, data))
        self.dataIsClean = False
        self.colFirst = (majDim == "COLS")
        self.manageWriteRate()
        results = self.sheets.service.spreadsheets().values().update(spreadsheetId=self.sheet, range=cellRange, body={ 'values': data, 'majorDimension': majDim}, valueInputOption="USER_ENTERED").execute()
        self.writeDataRangeCache(cellRange, data)

    def __cacheRow(self, row, data):
        """row is 1-indexed"""
        if self.colFirst:
            for c in range(len(data)):
                self.data[c][row-1] = data[c]
        else:
            data.extend(["" for _ in range(len(data), max([ len(r) for r in self.data ]))])
            self.data[row-1] = data
        return data
    
    def __cacheCol(self, col, colData):
        """col is 1-indexed"""
        if self.colFirst:
            self.data[col-1] = colData
        else:
            logger.debug(self.data)
            for r in range(len(colData)):
                logger.debug("{}, {}".format(r, colData[r]))
                self.data[r][col-1] = colData[r]
    
    def expandDataToCell(self, row, col):
        """0-indexed coordinates"""
        logger.debug("Expanding to: ({}, {})".format(row, col))
        if self.data == None:
            self.data = []
        if row < len(self.data) and col < max([len(r) for r in self.data]):
            logger.debug("Data size sufficient, not altering")
            return
        # Pad self.data to bottom right most cell location
        while len(self.data) < row + 1:
            self.data.append(["" for _ in range(col+1)])
        # ensure2D(self.data)
        for i, r in enumerate(self.data):
            if len(r) < col + 1:
                r = r + ["" for _ in range(col-len(r)+1)]
            self.data[i] = r
        logger.debug("Sized: {}".format(self.data))
        
    def writeDataRangeCache(self, cellRange, data):
        logger.debug("Caching: {}: {}".format(cellRange, data))
        if len(data) == 0:
            data = [[""]]
        if len(data) == 1 and len(data[0]) == 1:
            logger.debug("Single value (or empty) shunting off to writeDataCellCache")
            return self.writeDataCellCache(cellRange, data[0][0])
        # rows and cols are 1-indexed
        if ':' in cellRange:
            cellRangeStart, cellRangeEnd = cellRange.split(':')
            sCol, sRow = splitCR(cellRangeStart)
            eCol, eRow = splitCR(cellRangeEnd)
        else:
            sCol, sRow = splitCR(cellRange)
            eCol = sCol + 1
            eRow = sRow + 1
        # logger.debug("{}: ({},{}) -> ({},{})".format(cellRange, sRow, sCol, eRow, eCol))
        # logger.debug("{}".format(data))
        # if self.data:
        #     logger.debug("Currently {} x {}".format(len(self.data), max([len(r) for r in self.data])))
        # else:
        #     logger.debug("Currently self.data is None")
        # logger.debug("Before: {}".format(self.data))
        if self.colFirst:
            self.expandDataToCell(eCol-1, eRow-1)
        else:
            self.expandDataToCell(eRow-1, eCol-1)
            
            # Fill data
            rowOffset = sRow - 1
            colOffset = sCol - 1
            for r in range(eRow-sRow+1):
                for c in range(eCol-sCol+1):
                    logger.debug("Writing from ({},{}) to ({},{})".format(r, c, r + rowOffset, c + colOffset))
                    try:
                        self.data[r + rowOffset][c + colOffset] = data[r][c]
                    except IndexError as e:
                        logger.exception("Index out of range: self.data: {} x {}, [{}][{}]; data: {} x {}, [{}][{}]".format(len(self.data), max([len(r) for r in self.data]), r + rowOffset, c + colOffset, len(data), max([len(r) for r in data]), r, c))
                        logger.exception("\tcellRange: {}, rows: {}-{}, cols: {}-{}".format(cellRange, sRow, eRow, sCol, eCol))
                        raise e
        # logger.debug("After: {}".format(self.data))
        

    def writeDataCellCache(self, cell, data):
        col, row = splitCR(cell)
        return self.writeDataCellRCCache(row, col, data)

    def writeDataCellRC(self, row, col, data, formatted=False, dateAsString=True):
        """Row and col sent in must be 1-indexed, to correspond with sheet row and col."""
        cell = self.writeData(str(intToSheetsCol(col)) + str(row), [[data]])
        self.writeDataCellRCCache(row, col, data)
        return cell
        
    def writeDataCellRCCache(self, row, col, data):
        """1-indexed row col"""
        # Adjust for 0-indexed self.data
        row -= 1
        col -=1
        if self.colFirst:
            self.expandDataToCell(col, row)
            self.data[col][row] = data
        else:
            self.expandDataToCell(row, col)
            self.data[row][col] = data
        return data
    
    def appendRowData(self, data):
        """Assumes row-major, assumes sheet contains one table that includes cell A1."""
        # check to see if data is 2D list
        if self.colFirst:
            logger.error("appending rows to column major data NOT YET SUPPORTED")
            # TODO Report error better and handle exception
            return { 'values':data }
        data = ensure2D(data)
        self.dataIsClean = False
        self.manageWriteRate()
        results = self.sheets.service.spreadsheets().values().append(spreadsheetId=self.sheet, range="A1", body={ 'values': data }, valueInputOption="USER_ENTERED").execute()
        self.appendRowDataCache(self, data)
        self.dataIsClean = True
        return results

    def appendRowDataCache(self, data):
        """Assumes row major"""
        data = ensure2D(data)
        if self.colFirst:
            logger.error("appending rows to column major data NOT YET SUPPORTED")
        else:
            for row in data:
                self.data.append(row)
        self.data = unSparse2D(self.data)
    
    def batchUpdate(self, body):
        self.dataIsClean = False
        self.manageWriteRate()
        return self.sheets.service.spreadsheets().batchUpdate(spreadsheetId=self.sheet, body=body).execute()
    
    def deleteRow(self, row, end = -1):
        """Row is 1-based, not Python 0-based, be careful!"""
        if end == -1:
            end = row + 1
        logger.debug("Deleting rows %r-%r, as API indexes %r-%r" % (row, end, row-1, end-1))
        requests = [ { 'deleteRange': \
                        { "range": { \
                            "sheetId": 0, \
                            "startRowIndex": row-1, \
                            "endRowIndex": end-1, \
                        }, \
                        "shiftDimension": "ROWS"}\
                   } ]
        body = { 'requests': requests }
        logger.debug(body)
        result = self.batchUpdate(body)
        logger.debug(result) 
        self.deleteRowCache(row, end)
        return result
    
    def deleteAll(self):
        """Row is 1-based, not Python 0-based, be careful!"""
        if self.data == None or self.data == []:
            return
        
        if self.colFirst:
            rows = len(self.data[0])
        else:
            rows = len(self.data)
        
        logger.debug("Deleting ALL DATA AND FORMATTING: rows %r-%r, as API indexes %r-%r" % (row, end, row-1, end-1))
        requests = [ { 'deleteRange': \
                        { "range": { \
                            "sheetId": 0, \
                            "startRowIndex": 0, \
                            "endRowIndex": rows-1, \
                        }, \
                        "shiftDimension": "ROWS"}\
                   } ]
        body = { 'requests': requests }
        logger.debug(body)
        result = self.batchUpdate(body)
        logger.debug(result)
        self.data = []
        return result
        
    def deleteRowCache(self, row, end = -1):
        """1-indexed row and end"""
        row -= 1
        if end == -1:
            end = row + 1
        else:
            end -= 1
        if self.colFirst:
            for r in self.data:
                logger.debug("Deleting rows {} to {}".format(row, end))
                del r[row:end]
        else:
            for i in range(row, end):
                logger.debug("Deleting row {}".format(i))
                del self.data[i]
        
    def __fetchData(self, cellRange, vRO="UNFORMATTED_VALUE", dTRO="FORMATTED_STRING", majDim="ROWS"):
        """Fetches select data from Google Sheets online."""
        logger.debug("Fetching cellRange '{}'".format(cellRange))
        self.manageReadRate()
        self.colFirst = (majDim == "COLS")
        results = self.sheets.service.spreadsheets().values().get(spreadsheetId=self.sheet, range=cellRange, valueRenderOption=vRO, dateTimeRenderOption=dTRO, majorDimension=majDim).execute()
        logger.debug("Fetched: {}".format(results))
        try:
            data = unSparse2D(results['values'])
        except:
            data = []
        logger.debug("Fetched: {}".format(data))
        # The cellRange requested is often truncated, adjust based on *actual* data returned.
        # If it's a single cell descriptor, then just leave it be
        if ':' in cellRange:
            crStart, crEnd = cellRange.split(':')
            logger.debug("{} : {}".format(crStart, crEnd))
            sCol, sRow = splitCR(crStart)
            eCol, eRow = splitCR(crEnd)
            logger.debug("{}: ({},{}) -> ({},{})".format(cellRange, sRow, sCol, eRow, eCol))
            # If no data returned, fake it as best we can.
            # This is usually a blank row or column
            if data == []:
                logger.debug("No data fetched, fixing sizes and filling in blank data")
                if self.colFirst:
                    eRow = max([len(c) for c in self.data])
                    eCol = sCol
                    data = [ [ "" for _ in range(eRow) ] ]
                else:
                    eRow = sRow
                    eCol = max([len(r) for r in self.data])
                    data = [ [ "" for _ in range(eCol) ] ]
            else:
                if self.colFirst:
                    eRow = sRow + max([len(c) for c in data]) - 1
                    eCol = sCol + len(data) - 1
                else:
                    eRow = sRow + len(data) - 1
                    eCol = sCol + max([len(r) for r in data]) - 1
            logger.debug("{}: ({},{}) -> ({},{})".format(cellRange, sRow, sCol, eRow, eCol))
            logger.debug("Data: {}".format(data))
            cellRange = intToSheetsCol(sCol) + str(sRow) + ":" + intToSheetsCol(eCol) + str(eRow)
        logger.debug('Caching cell range: {}'.format(cellRange))
        if data == []:
            data = [[""]]
        self.writeDataRangeCache(cellRange, data)
        return data

    def __fetchAllData(self, vRO="UNFORMATTED_VALUE", dTRO="FORMATTED_STRING", majDim="ROWS"):
        """Fetches all data from Google Sheets online."""
        self.manageReadRate()
        self.colFirst = (majDim == "COLS")
        results = self.sheets.service.spreadsheets().values().get(spreadsheetId=self.sheet, range="A1:ZZ500", valueRenderOption=vRO, dateTimeRenderOption=dTRO, majorDimension=majDim).execute()
        try:
            data = unSparse2D(results['values'])
        except:
            data = [[]]
        self.data = data
        logger.debug("Fetched all data:\n{}".format(self.data))
        return data
        
    def getDataRange(self, cellRange, formatted=False, dateAsString=True, colFirst=False):
        """Wraps call to __fetchData, converting bool flags to Google Sheets strings."""
        if not self.data or not self.dataIsClean:
            data = self.__fetchData(cellRange, vRO="FORMATTED_VALUE" if formatted else "UNFORMATTED_VALUE", dTRO="FORMATTED_STRING" if dateAsString else "SERIAL_NUMBER", majDim="COLUMNS" if colFirst else "ROWS")
        else:
            cellRangeStart, cellRangeEnd = cellRange.split(':')
            sCol, sRow = splitCR(cellRangeStart)
            eCol, eRow = splitCR(cellRangeEnd)
            data = []
            if self.colFirst:
                logger.error("Not yet implemented for colFirst mode")
            else:
                for r in range(sRow, eRow+1):
                    data.append(self.data[sCol:eCol+1])
        return data
    
    def getAllData(self, formatted=False, dateAsString=True, colFirst=False):
        """Wraps call to __fetchAllData, converting bool flags to Google Sheets strings."""
        vRO = "FORMATTED_VALUE" if formatted else "UNFORMATTED_VALUE"
        dTRO = "FORMATTED_STRING" if dateAsString else "SERIAL_NUMBER"
        majDim = "COLUMNS" if colFirst else "ROWS"
        logger.debug("self.__fetchAllData(vRO={}, dTRO={}, majDim={})".format(vRO, dTRO, majDim))
        results = self.__fetchAllData(vRO=vRO, dTRO=dTRO, majDim=majDim)
        # logger.debug("{}".format(self.data))
        return self.data
        
    def getCell(self, cell, formatted=False, dateAsString=True):
        """1-indexed Xn cell"""
        col,row = splitCR(cell)
        return self.getCellRC(row, col, formatted, dateAsString)

    def getCellCache(self, cell):
        """1-indexed Xn cell"""
        col,row = splitCR(cell)
        return self.getCellRCCache(row, col)
        
    
    def getCellRC(self, row, col, formatted=False, dateAsString=True):
        """Row and col are 1-indexed"""
        if not self.data or not self.dataIsClean:
            data = self.__fetchData(str(intToSheetsCol(col)) + str(row), vRO="FORMATTED_VALUE" if formatted else "UNFORMATTED_VALUE", dTRO="FORMATTED_STRING" if dateAsString else "SERIAL_NUMBER")[0][0]
            self.writeDataCellRCCache(row, col, data)
        else:
            data = self.getCellRCCache(row, col)
        return data
        
    def getCellRCCache(self, row, col):
        """Row and col are 1-indexed args, 0-indexed cached data"""
        row -= 1
        col -= 1
        try:
            if self.colFirst:
                data = self.data[col][row]
            else:
                data = self.data[row][col]
        except IndexError as e:
            logger.debug("Requested data from a cell not in sheet range, returning None")
            data = None
        return data
    
    def getRow(self, row, formatted=False, dateAsString=True):
        """Return the 1-indexed row.  If row is blank, list is empty, regardless of number of columns in data?"""
        logger.debug("Getting 1-indexed row %d" % (row))
        if self.data == None:
            self.getAllData()
        if not self.dataIsClean:
            logger.debug("Data dirty, re-fetching")
            rowData = self.__fetchData("A" + str(row) + ":AA" + str(row), vRO="FORMATTED_VALUE" if formatted else "UNFORMATTED_VALUE", dTRO="FORMATTED_STRING" if dateAsString else "SERIAL_NUMBER")[0]
            logger.debug("Fetched: %r" % (rowData))
            rowData = self.__cacheRow(row, rowData)
        else:
            rowData = self.getRowCache(row)
        logger.debug("got row {}".format(row))
        return rowData
    
    def getRowCache(self, row):
        if self.colFirst:
            rowData = []
            for c in range(len(self.data)):
                rowData.append(self.data[c][row-1])
        else:
            rowData = self.data[row-1]
        return rowData

    def getCol(self, col, formatted=False, dateAsString=True):
        """Return the 1-indexed column"""
        colAlpha = intToSheetsCol(col)
        logger.debug("Getting 1-indexed column %d (%s)" % (col, colAlpha))
        if self.data == None:
            self.getAllData()
        if not self.dataIsClean:
            logger.debug("Data dirty, re-fetching")
            colData = self.__fetchData(colAlpha + "1:" + colAlpha + "1000", vRO="FORMATTED_VALUE" if formatted else "UNFORMATTED_VALUE", dTRO="FORMATTED_STRING" if dateAsString else "SERIAL_NUMBER")
            colData = [ d[0] for d in colData ]
            logger.debug("Fetched: %r" % (colData))
            self.__cacheCol(col, colData)
        else:
            colData = self.getColCache(col)
        return colData

    def getColCache(self, col):
        logger.debug("Data clean, returning cached data")
        if self.colFirst:
            colData = self.data[col-1]
        else:
            colData = []
            for r in range(len(self.data)):
                colData.append(self.data[r][col-1])
        return colData

class GoogleSheetCached(GoogleSheetAtomic):
    """Assume a complete lock of data on the google side while operating on it.  i.e. no one is changing anything on the website or from another process.  Much faster, but riskier."""
    def __init__(self, sheetsAccessor, name="", sheetID="", createIfAbsent=False):
        super().__init__(sheetsAccessor, name, sheetID, createIfAbsent)
        if not self.sheet:
            raise NoSheetFile(name)
        # Get the data *once* when we start
        self.getAllData()
    
    def setupDefault(self):
        pass
    
    def __clearOnline(self):
        """Clear everything online, including formatting.  DOES NOT TOUCH self.data CACHE"""
        if self.data == None or self.data == []:
            return
    
        if self.colFirst:
            rows = len(self.data[0])
        else:
            rows = len(self.data)
    
        logger.debug("Deleting ALL DATA AND FORMATTING: rows 1-%r, as API indexes 0-%r" % (rows, rows-1))
        requests = [ { 'deleteRange': \
                        { "range": { \
                            "sheetId": 0, \
                            "startRowIndex": 0, \
                            "endRowIndex": rows-1, \
                        }, \
                        "shiftDimension": "ROWS"}\
                   } ]
        body = { 'requests': requests }
        logger.debug(body)
        result = self.batchUpdate(body)
        logger.debug(result)
        return result
        
    def pushOnline(self):
        numCols = max( [ len(row) for row in self.data] )
        numRows = len(self.data)
        if self.colFirst:
            temp = numCols
            numCols = numRows
            numRows = temp
        self.__clearOnline()
        super().writeData("A1:" + intToSheetsCol(numCols) + str(numRows), self.data)
    
    def getCol(self, col):
        return self.getColCache(col)
    
    def getRow(self, row):
        return self.getRowCache(row)
    
    def getCell(self, cellRange):
        return self.getCellCache(cellRange)
    
    def getCellRC(self, row, col):
        return self.getCellRCCache(row, col)
    
    def deleteRow(self, row):
        return self.deleteRowCache(row)
    
    def deleteAll(self):
        self.data = []
        return self.data
    
    def appendRowData(self, data):
        return self.appendRowDataCache(data)
    
    def writeDataCellRC(self, row, col, data):
        return self.writeDataCellRCCache(row, col, data)
    
    def writeDataCell(self, cellRange, data):
        return self.writeDataCellCache(cellRange, data)
    
    def insertBlankRowsAfter(self, row, end):
        return self.insertBlankRowsAfterCache(row, end)
    
    def sortOnCol(self, col):
        return self.sortOnColCache(col)
    
    def writeData(self, cellRange, data, majDim="ROWS"):
        return self.writeDataRangeCache(cellRange, data)
    
    
    
# Hiding inside a class, because unittest only searches for module-level classes.  This prevents the inner class from being treated as a runnable TestCase.  Crude, but effective!
class TestBase:
    class TestGoogleSheetsBase(unittest.TestCase):
        def __init__(self, argv):
            super().__init__(argv)
            self.maxDiff = None

        initialTestData = [ \
            [ "Date", "Calendar", "IDs", "Description", "Start", "End", "Location"], \
            [ "4/18/2019", "Cal A", "1,2", "Foo", "11:00:00 AM", "2:00:00 PM", "Staff Office"], \
            [ "4/19/2019", "Cal B", "3", "Bar", "9:00:00 AM", "5:00:00 PM", "" ], \
            [ "4/20/2019", "Cal A", "4", "Goo", "9:00:00 AM", "2:00:00 PM", "Unicorn Meadows" ], \
            [ "4/21/2019", "Cal A", "5", "Car", "12:00:00 AM", "12:00:00 AM", "" ], \
            [ "4/23/2019", "Cal B", "6, 7", "Moo", "3:00:00 AM", "5:00:00 PM", "" ] \
        ]

        @classmethod
        def setUpClass(cls):
            TestBase.TestGoogleSheetsBase.sheet = None
        
        ### self.sheet and self.initialTestData are set in setUp() routine in subclasses.
                
        def setUp(self):
            logger.debug("----------")
            self.sheet = TestBase.TestGoogleSheetsBase.sheet
            self.initialTestData = TestBase.TestGoogleSheetsBase.initialTestData

        def test_00_writeData(self):
            self.sheet.writeData("A1:G6", self.initialTestData)
            self.assertEqual(self.sheet.data, self.initialTestData)
        
        def test_01_makeHeaderRow(self):
            self.sheet.makeHeaderRow()
        
        def test_02_makeBigBoldText(self):
            self.sheet.makeBigBoldText("A1:G1")
    
        def test_getCol(self):
            self.assertEqual(self.sheet.getCol(2), [ r[1] for r in self.initialTestData ])
        
        def test_getDataRange(self):
            self.assertEqual(self.sheet.getDataRange("B2:D3"), [["Cal A", "1,2", "Foo"], ["Cal B", "3", "Bar"]])
        
        def test_getRow(self):
            self.assertEqual(self.sheet.getRow(3), self.initialTestData[2])
    
        def test_getCellRC(self):
            self.assertEqual(self.sheet.getCellRC(3, 3), self.initialTestData[2][2])
        
        def test_writeDataOneCell(self):
            self.sheet.writeData("A1", [["A"]])
            self.assertEqual(self.sheet.data[0][0], "A")
            self.sheet.writeData("A1", [[self.initialTestData[0][0]]])

            self.sheet.writeData("G1", [["Z"]])
            self.assertEqual(self.sheet.data[0][6], "Z")
            self.sheet.writeData("G1", [[self.initialTestData[0][6]]])
        
        def test_writeDataRange(self):
            self.sheet.writeData("B1:B3", [["A"], ["B"], ["C"]])
            self.assertEqual(self.sheet.getCol(2)[0:3], ["A", "B", "C"])
            self.assertEqual(self.sheet.data[0][1], "A")
            self.assertEqual(self.sheet.data[1][1], "B")
            self.assertEqual(self.sheet.data[2][1], "C")
            self.sheet.writeData("B1:B3", [[self.initialTestData[0][1]], [self.initialTestData[1][1]], [self.initialTestData[2][1]]])
        
        def test_insertBlankRowAfter_deleteRow(self):
            if self.sheet.colFirst:
                beforeSize = 0 if not self.sheet.data else len(self.sheet.data[0])
            else:
                beforeSize = 0 if not self.sheet.data else len(self.sheet.data)
            self.sheet.insertBlankRowAfter(1)
            self.assertEqual(self.sheet.getRow(2), ["", "", "", "", "", "", ""])
            self.assertEqual(self.sheet.data[1], ["", "", "", "", "", "", ""])
            if self.sheet.colFirst:
                afterSize = len(self.sheet.data[0])
            else:
                afterSize = len(self.sheet.data)
            self.assertEqual(beforeSize + 1, afterSize)

            self.sheet.deleteRow(2)
            self.assertEqual(self.sheet.getRow(2), self.initialTestData[1])
            self.assertEqual(self.sheet.data[1], self.initialTestData[1])
            if self.sheet.colFirst:
                afterSize = len(self.sheet.data[0])
            else:
                afterSize = len(self.sheet.data)
            self.assertEqual(beforeSize, afterSize)
    
        def test_intToSheetsCol(self):
            self.assertEqual(intToSheetsCol(1), 'A')
            self.assertEqual(intToSheetsCol(26), 'Z')
            self.assertEqual(intToSheetsCol(27), 'AA')
            self.assertEqual(intToSheetsCol(52), 'AZ')
            self.assertEqual(intToSheetsCol(53), 'BA')

        def test_sheetsColToInt(self):
            self.assertEqual(sheetsColToInt('A'), 1)
            self.assertEqual(sheetsColToInt('Z'), 26)
            self.assertEqual(sheetsColToInt('AA'), 27)
            self.assertEqual(sheetsColToInt('AZ'), 52)
            self.assertEqual(sheetsColToInt('BA'), 53)
    
        def test_sortOnCol(self):
            self.sheet.sortOnCol(1)
            self.assertEqual(self.sheet.getCol(2), 
                [self.initialTestData[0][1]] + sorted([ r[1] for r in self.initialTestData[1:] ]) )
            self.sheet.sortOnCol(0)
        
        def test_splitCR(self):
            self.assertEqual(splitCR('A1'), (1, 1))
            self.assertEqual(splitCR('AZ50'), (52, 50))
    
        def test_ensure2D(self):
            self.assertEqual(ensure2D(1), [[1]])
            self.assertEqual(ensure2D([]), [[]])
            self.assertEqual(ensure2D([[]]), [[]])
    
        def test_unSparse2D(self):
            self.assertEqual(unSparse2D([["1"], ["2", "3"]]), [["1", ""], ["2", "3"]])
    
        def test_trimList(self):
            self.assertEqual(trimList([1, 2, 3, '']), [1, 2, 3])
            self.assertEqual(trimList([1, 2, '', 3, '']), [1, 2, '', 3])
            self.assertEqual(trimList(['']), [])
            self.assertEqual(trimList([]), [])
        

class TestGoogleSheets(TestBase.TestGoogleSheetsBase):
    def __init__(self, argv):
        super().__init__(argv)
        self.maxDiff = None

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        sheets = GoogleSheets(os.path.abspath('client_secret-omg-events.json'), os.path.abspath('credentials.json'), "omg-events@omg.org")
        try:
            TestBase.TestGoogleSheetsBase.sheet = GoogleSheetAtomic(sheets, os.path.join(["configs","anytown-19"]))
        except NoSheetFile as e:
            logger.exception(e)
            exit(-1)



class TestGoogleSheetsCached(TestBase.TestGoogleSheetsBase):
    """Same exact tests as the Atomic version, just using the Cached version"""
    def __init__(self, argv):
        super(TestGoogleSheetsCached, self).__init__(argv)

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        sheets = GoogleSheets(os.path.abspath('client_secret-omg-events.json'), os.path.abspath('credentials.json'), "omg-events@omg.org")
        try:
            TestBase.TestGoogleSheetsBase.sheet = GoogleSheetCached(sheets, os.path.join(["configs","anytown-19"]))
        except NoSheetFile as e:
            logger.exception(e)
            exit(-1)
        # Get all data once, prior to running tests
        TestBase.TestGoogleSheetsBase.sheet.getAllData()
    
    

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description = 'Google Sheets Accessor')
    # Debugging and testing
    parser.add_argument('--verbose', default=False, action='store_true')
    parser.add_argument('--debug', default=False, action='store_true')
    parser.add_argument('--test')
    args = parser.parse_args()

    # Logging to file is always set to debug, this is just for console
    if args.verbose:
        logger.setLevel(erlogging.INFO)
    if args.debug:
        logger.setLevel(erlogging.DEBUG)
    
    testArgs = [os.path.basename(__file__)]
    if args.test:
        testArgs.append(args.test)
    
    gs = GoogleSheets(os.path.abspath('client_secret-omg-events.json'), os.path.abspath('credentials.json'), "omg-events@omg.org")
    
    # gs.getSheets()
    
    unittest.main(argv=testArgs)
