from __future__ import print_function
#googleapiclient.discovery
from apiclient.discovery  import build
from httplib2 import Http
from oauth2client import file, client, tools
from oauth2client.contrib import gce
from apiclient.http import MediaFileUpload
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import os
import pathlib
from pathlib import Path
import io
import glob
import itertools
import shutil
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import (
    LineChart,
    Reference,
    Series
)

import random
import time
from time import sleep
import datetime
import csv


CLIENT_SECRET = "/Users/eyana.mallari/Projects-Local/client_secret_eyanag.json"



# --------------------------------
# GDrive API: GDrive Authorization
# --------------------------------

SCOPES='https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
    creds = tools.run_flow(flow, store)
DRIVE_SERVICE = build('drive', 'v3', http=creds.authorize(Http()))
SS_SERVICE = build('sheets', 'v4', http=creds.authorize(Http()))


FOLDER_ID = '19FBo4iSjyCS3NcqX6zaNgxtXWMqW7AyM'
SS_ID = '1qs2GXzxhbn8klcB3Y-qQVJs0CRt8iQ20LyntKP36BRY'
SHEET_ID=227869185
TAB_NAME ='Folders Lookup!A2:H'
now = datetime.datetime.now()

def countFiles(folder_id):
# Main Folder

    results = DRIVE_SERVICE.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and parents in '"+folder_id+"' and trashed = false",fields="nextPageToken, files(id, name)").execute()

    items = results.get('files', [])

    count = 0
    for item in items:
        #print('{0} ({1})'.format(item['name'], item['id']))
        count += 1
    return count

def getLatestFile(folder_id):
# Main Folder

    results = DRIVE_SERVICE.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and parents in '"+folder_id+"' and trashed = false",fields="nextPageToken, files(id, name, createdTime)", orderBy="createdTime").execute()


    #results = DRIVE_SERVICE.files().list(q="parents in '"+folder_id+"' and trashed = false",fields="nextPageToken, files(id, name, createdTime)", orderBy="createdTime").execute()


    items = results.get('files', [])

    if not items:
        return ""
    else:
        return items
def retrieveId(items,count):
    if not items:
        return ""
    else:
        print(items[count-1]['id'])
        return items[count-1]['id']


def retrieveName(items,count):
    if not items:
        return ""
    else:
        print(items[count-1]['name'])
        return items[count-1]['name']

def retrieveCreatedDate(items,count):
    if not items:
        return ""
    else:
        print(items[count-1]['createdTime'])
        return items[count-1]['createdTime']



def writeToSheets(values,range,spreadsheet_id):
    print("Writing to Sheets")
    body = {
        'values': values
        }
    print(values)
    SS_SERVICE.spreadsheets().values().update(spreadsheetId=spreadsheet_id,valueInputOption='USER_ENTERED',range=range,body=body).execute()


# -------------------------------------
# GDrive API: Retrieves all Folders
# -------------------------------------
def getAllFolders(folder_id, spreadsheet_id):
# Main Folder
    folder_id = FOLDER_ID

    results = DRIVE_SERVICE.files().list(q="mimeType='application/vnd.google-apps.folder' and parents in '"+folder_id+"' and trashed = false",fields="nextPageToken, files(id, name)",pageSize=400).execute()

    items = results.get('files', [])

    values = []
    #print(type(items))
    #print(items)

    count = 1
    print(len(items))
    for item in items:
        print(count)
        print('{0} ({1})'.format(item['name'], item['id']))

        isManager = item['name'].find("] ")+1

        if isManager > 1:
            person_name = item['name'][isManager+15:].strip()
        else:
            person_name = item['name'][5:].strip()


        file_count = countFiles(item['id'])
        latestFileId = '=HYPERLINK("https://docs.google.com/spreadsheets/d/'+retrieveId(getLatestFile(item['id']),file_count)+'","Click")'

        latestFile = retrieveName(getLatestFile(item['id']),file_count)
        createdDate = retrieveCreatedDate(getLatestFile(item['id']),file_count)

        print("----", person_name)
        print("----", file_count)
        print("----", createdDate)

        item['person'] = person_name
        item['file_count'] = file_count
        #item['link'] ="https://drive.google.com/drive/folders/"+item['id']
        item['link'] = '=HYPERLINK("https://drive.google.com/drive/folders/'+item['id']+'","Click")'

        values.append([item['person'],item['name'],item['link'],item['id'],item['file_count'],latestFileId,createdDate,latestFile])


        count+=1
        #range_ = 'Folder Lookup June 8!A'+str(count)+':F'+str(count)
        #values = [[item['person'],item['name'],item['link'],item['id'],item['file_count'],str(now)]]
    writeToSheets(values,TAB_NAME,spreadsheet_id)





def colorAndSort(spreadsheet_id):
    sheet_id=SHEET_ID
    my_range = {
        'sheetId': sheet_id,
        'startRowIndex': 1,
        'endRowIndex': 600,
        'startColumnIndex': 5,
        'endColumnIndex': 5,
    }
    requests = [{
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [ my_range ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [ { 'userEnteredValue': '=GT($E1,0)'} ]
                    },
                    'format': {
                        'backgroundColor': { 'red': 1, 'green': 1, 'blue': 0 }
                        }
                    }
                  }
                }
            },
        {
        "sortRange": {
            "range": {
            "sheetId": sheet_id,
            "startRowIndex": 1,
            "endRowIndex": 600,
            "startColumnIndex": 0,
            "endColumnIndex": 7
        },
        "sortSpecs": [
          {
            "dimensionIndex": 4,
            "sortOrder": "DESCENDING"
          },
          {
            "dimensionIndex": 1,
            "sortOrder": "ASCENDING"
          }
        ]
      }
    }
          ]

    body = {
        'requests': requests
    }
    response = SS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} cells updated.'.format(len(response.get('replies'))));


def readSheet():

    SPREADSHEET_ID =  SS_ID
    RANGE_NAME = 'Test!A1:B'
    results = SS_SERVICE.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                                 range=RANGE_NAME).execute()
    values = results.get('values', [])
    if not values:
        print('No data found.')
    else:
        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))


def main():
    getAllFolders(FOLDER_ID,SS_ID)
    colorAndSort(SS_ID)
    #loopRosterCreateFiles(getSalesRep())

if __name__ == '__main__':
    main()
