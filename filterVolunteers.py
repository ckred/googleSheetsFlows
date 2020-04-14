#!/usr/local/opt/python/bin/python3.7

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
ALL_VOLUNTEERS_SPREADSHEET = '18m_MDb3aqmKvZRndn34077hETa3rxskeb0WTyR6jKlQ'
ALL_VOLUNTEERS_RANGE = 'Form Responses 1!A2:H'
PENPALS_SPREADSHEET = '1_Oh_rqOyCJo9LhDX4OZKqde18jHsroRAXQZRx6MN84c'
PENPALS_READ_RANGE = 'Pen Pals List!B2:B'
PENPALS_UPDATE_RANGE = 'Pen Pals List!B2:E'
FEED_SPREADSHEET = '18m_MDb3aqmKvZRndn34077hETa3rxskeb0WTyR6jKlQ'
FEED_READ_RANGE = 'FEED Volunteers!B2:B'
FEED_UPDATE_RANGE = 'FEED Volunteers!A2:D'
PATH_TO_CREDENTIALS = '/Users/caryl/Dev/wtcfSheets/credentials.json'

def main():
    """
    This application retrieves the latest list from the source spreadsheet (in our case, ALL_VOLUNTEERS), 
    and then filters them out conditionally based on the contents of the 6th column (which for us contains the 
    projects of interest). Based on the values, data from each row may be copied to 'child' tabs or spreadsheets.
    This is designed to be run as a batch job to keep the downstream sheets in eventual sync (one-directional) with
    the source sheet.
    """   
    service = getService()
    
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ALL_VOLUNTEERS_SPREADSHEET,
                                range=ALL_VOLUNTEERS_RANGE).execute()
    values = result.get('values', [])

    penpals = currentList(PENPALS_SPREADSHEET,PENPALS_READ_RANGE)
    feed = currentList(FEED_SPREADSHEET, FEED_READ_RANGE)

    if not values:
        print('No data found.')
    else:
        for row in values: 
            if(('FEED' in row[5]) and (row[2] not in feed)): # If this volunteer is interested in the FEED project
                print(feed)
                print(row[2] + " added to FEED volunteers.") # Add their data to the spreadsheet of FEED volunteers
                addVolunteer([[row[1],row[2],row[3],row[4]]],FEED_SPREADSHEET,FEED_UPDATE_RANGE)
                # break; # for testing purposes, allows you to add one person at a time
            else:
                continue
            if(('Pen Pal' in row[5]) and (row[2] not in penpals)): # If this volunteer is interested in the Pen Pals project
                print(row[2] + " added to Pen Pal volunteers.")
                addVolunteer([['',row[2],'A',row[4],row[1]]],PENPALS_SPREADSHEET,PENPALS_UPDATE_RANGE) # Add them to the Pen Pals spreadsheet
            else:
                continue

def addVolunteer(values,inputSpreadsheet,inputRange):
    """
    Given a list of relevant values (in the expected order), the spreadsheet where the values should be added, and the 
    range where they need to be added, this function will add another row to the destination spreadsheet with the specified
    list of values, starting with the first column of the tab (not the range).
    """
    service = getService()
    
    body = {
        'values': values
    }
    value_input_option = 'USER_ENTERED'
    result = service.spreadsheets().values().append(
        spreadsheetId=inputSpreadsheet, range=inputRange, body=body, valueInputOption=value_input_option).execute()
    print('{0} cells appended.'.format(result \
                                           .get('updates') \
                                           .get('updatedCells')))

def currentList(inputSpreadsheet,inputRange):
    """
    Given a source sheet, this will read all of the values in the specified range and return an array containing those values.
    """
    service = getService()

    readSheet = service.spreadsheets()
    result = readSheet.values().get(spreadsheetId=inputSpreadsheet,
                                range=inputRange).execute()
    volunteerList = result.get('values', [])
    volunteers = []
    for sublist in volunteerList:
        for volunteer in sublist:
            volunteers.append(volunteer)
    return volunteers;

def getService():
    """
    This authenticates into the Google Sheets API assuming there is a credentials.json in the path specified in the 
    PATH_TO_CREDENTIALS global variable.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                PATH_TO_CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service;

if __name__ == '__main__':
    main()
