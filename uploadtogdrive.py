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



FILE_MASTERLIST = 'Input/CONTACTS_ALL_STATIC.xlsx'
FILE_TEMPLATE ='Input/Contacts_Template.xlsx'
OUTPUT_DIRECTORY = 'Output/'
EXCLUDE_LIST = []
INCLUDE_LIST = ['Calista Rosales', 'Bianca Cardenas', 'Colette Black']

now = datetime.datetime.now()
QUARTER = "2019Q2"
MONTH = "May 2019"
VERSION = str(now.month).zfill(2) + str(now.day).zfill(2)

df = pd.read_excel(FILE_MASTERLIST,"CONTACTS_FINAL")


# --------------------------------
# GDrive API: GDrive Authorization
# --------------------------------

SCOPES='https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
    creds = tools.run_flow(flow, store)
SERVICE = build('drive', 'v3', http=creds.authorize(Http()))
SS_SERVICE = build('sheets', 'v4', http=creds.authorize(Http()))


PARENT_FOLDER = '19FBo4iSjyCS3NcqX6zaNgxtXWMqW7AyM'


# ------------------------------------
# GDrive API: Check if Filename exists
# ------------------------------------
def fileInGDrive(filename):
    results = SERVICE.files().list(q="mimeType='application/vnd.google-apps.spreadsheet' and name='"+filename+"' and trashed = false and parents in '"+PARENT_FOLDER+"'",fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if items:
        return True
    else:
        return False

# ------------------------------------
# GDrive API: Check if Folder exists
# ------------------------------------
def folderInGDrive(filename):
    results = SERVICE.files().list(q="mimeType='application/vnd.google-apps.folder' and name='"+filename+"' and trashed = false and parents in '"+PARENT_FOLDER+"'",fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if items:
        return True
    else:
        return False

# ---------------------------------------
# GDrive API: Create New Folder
# ---------------------------------------
def createGDriveFolder(filename,parent):
    file_metadata = {'name': filename,'parents': [parent],
    'mimeType': "application/vnd.google-apps.folder"}

    folder = SERVICE.files().create(body=file_metadata,
                                        fields='id').execute()
    print('Upload Success!')
    print('FolderID:', folder.get('id'))
    return folder.get('id')


# ---------------------------------------
# GDrive API: Upload files to Google Drive
# ---------------------------------------
def writeToGDrive(filename,source,folder_id):
    file_metadata = {'name': filename,'parents': [folder_id],
    'mimeType': 'application/vnd.google-apps.spreadsheet'}
    media = MediaFileUpload(source,
                            mimetype='application/vnd.ms-excel')

    if fileInGDrive(filename) is False:
        file = SERVICE.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        print('Upload Success!')
        print('File ID:', file.get('id'))
        return file.get('id')

    else:
        print('File already exists as', filename)


# ---------------------------------------
# GSheet API: Freeze Cells
# ---------------------------------------
def freezeCells(spreadsheet_id,sheet_id):

    my_range = {
        'sheetId': sheet_id
    }
    requests = [{
            "updateSheetProperties": {
                'properties': {
                    'sheetId': sheet_id,
                    'gridProperties': { 'frozenRowCount': 1,'frozenColumnCount': 2}
                },
                'fields': 'gridProperties(frozenRowCount,frozenColumnCount)'
            }
        }
    ]

    body = {
        'requests': requests
    }
    response = SS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} update.'.format(len(response.get('replies'))));


# ---------------------------------------
# GSheet API: Protect Cells
# ---------------------------------------
def protectCells(spreadsheet_id,sheet_id):

    my_range = {
        'sheetId': sheet_id
    }
    requests = [{
      "addProtectedRange": {
        "protectedRange": {
          "range": {
                "sheetId": sheet_id,
                "startColumnIndex": 0,
                "endColumnIndex": 12,
            },
         "editors": {
            "users": [
              "eyanamallari@gmail.com"
            ]
          }

          }
     }
    }

    ]

    body = {
        'requests': requests
    }
    response = SS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} update.'.format(len(response.get('replies'))));


# ---------------------------------------
# GSheet API: Delete unnecessary cells
# ---------------------------------------
def deleteCells(spreadsheet_id,sheet_id):

    my_range = {
        'sheetId': sheet_id
    }
    requests = [
    {
      "deleteDimension": {
        "range": {
          "sheetId": sheet_id,
          "dimension": "COLUMNS",
          "startIndex": 8,
          "endIndex": 26
        }
      }
    }
    ]

    body = {
        'requests': requests
    }
    response = SS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} update.'.format(len(response.get('replies'))));

# ---------------------------------------
# GSheet API: Loop
# ---------------------------------------

def loopGSpreadsheet(spreadsheet_id):
    sheet_metadata = SS_SERVICE.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    sheets = sheet_metadata.get('sheets', '')

    for sheet in sheets:
        sheet_title = sheet.get("properties", {}).get("title")
        print("Processing ", sheet_title)
        #print(sheet)

        sheet_id = sheet.get("properties", {}).get("sheetId")
        freezeCells(spreadsheet_id,sheet_id)
        #protectCells(spreadsheet_id,sheet_id)

        column_count = sheet.get("properties", {}).get("gridProperties").get("columnCount")
        print(column_count, "columns")

        if column_count > 8:
            deleteCells(spreadsheet_id,sheet_id)




def getFolder(person):
    with open(FILE_FOLDERS, mode='r') as infile:
        reader = csv.reader(infile)
        folders_lookup = {rows[0]:rows[3]for rows in reader}

    person_string = str(person).strip()
    if person_string in folders_lookup:
        folder_id = folders_lookup.get(person_string)
        return folder_id
    else:
        return "No Folder"
        print('FOLDER NOT FOUND!!!!!', person)


def getFolderfromGDrive(folder_name):
# Main Folder

    results = SERVICE.files().list(q="mimeType='application/vnd.google-apps.folder' and name='"+folder_name+"' and trashed = false and parents in '"+PARENT_FOLDER+"'",fields="nextPageToken, files(id, name)").execute()

    items = results.get('files', [])

    #print(items)

    if not items:
        return ""
    else:
        print(items[-1]['name'])
        return items[-1]['id']


# ---------------------------------------
# Upload Files in GDrive
# ---------------------------------------
def loopRosterUploadFiles(reps):

    print("Creating files!")
    print(reps)

    count = 1

    for rep in reps:
        print('------')

        print(count, rep)

        count = count + 1

        output_rep = "Contacts List - " + str(rep)
        output_folder = OUTPUT_DIRECTORY+ output_rep

        # Check if Folder for rep exists
        if not os.path.exists(output_folder):
            print("Folder not found for", rep)
        print("Created folder for", rep)

        # Set File Names
        rep_excel_file = QUARTER+ " Contact List - "+str(rep)+ " " + VERSION +".xlsx".strip()
        rep_excel_file_no_ext = QUARTER+ " Contact List - "+str(rep)+ " " + VERSION.strip()
        rep_excel_path = output_folder+"/"+rep_excel_file


        # Create folder if it doesnt exist yet
        if folderInGDrive(output_rep) is False:
            createGDriveFolder(output_rep,PARENT_FOLDER)
            print('Folder created for', rep)

        loopGSpreadsheet(writeToGDrive(rep_excel_file_no_ext,rep_excel_path,getFolderfromGDrive(output_rep)))



def getSalesRep():
    df['Sales Representative'].fillna('Unknown', inplace = True)
    print('Getting all Sales Representatives')

    # --- EXCLUDE REPS ----
    #df_filtered = df[~df['Sales Representative'].isin(EXCLUDE_LIST)]
    #df_roster = df_filtered['Sales Representative'].unique()


    # --- INCLUDE REPS ----
    df_filtered = df[df['Sales Representative'].isin(INCLUDE_LIST)]
    df_roster = df_filtered['Sales Representative'].unique()

    #df_roster = df['Sales Representative'].unique()

    print(df_roster)
    d = df_roster.tolist()
    d_final = d

    #d_final = {k: d[k]for k in list(d)}

    print(d)
    print(len(d))

    print(d_final)
    print(len(d_final))
    return d_final




def generateNewFolders(reps):

    # FOLDERS LOOKUP
    count = 1
    folders = {}
    managers =[]
    for rep in reps:
        foldername = "Contacts List - " + str(rep)
        if folderInGDrive(foldername) == False:
            folder_id = createGDriveFolder(foldername,PARENT_FOLDER)
            print("New Folder", foldername, folder_id)
            count += 1
        # else:
        #     print(foldername, "exists!")
        time.sleep(5)
    print("Created", count, "new folders")






def main():
    #generateNewFolders(getSalesRep())
    loopRosterUploadFiles(INCLUDE_LIST)


if __name__ == '__main__':
    main()
