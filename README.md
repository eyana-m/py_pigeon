# Distribute data using python

* Split and distribute **files** to **sales representative** automatically

## Tools:
1. Python
2. Google API (GDrive, GSheet)

## Tutorials

1. Split Excel files - [Click here](http://eyana.me/split-excel-files-using-python/)
2. Upload files to Google Drive - [Part 1](http://eyana.me/upload-files-to-gdrive-using-python-part-1/), [Part 2]((http://eyana.me/upload-files-to-gdrive-using-python-part-1/))
4. Upload to GDrive - Python, GDrive API
5. Freeze top visible row and column - Python, GSheet API



# Workflow

1. Generate new GDrive folders for Teamlink connections without folders
2. Save folders lookup as Input
3. Split the CTD list by connections. Save to respective GDrive folders
4. Save contacts lookup as Input
5. Send emails via Python

## Input:

### For Part 1 - Split:
* List of contacts with the following fields:
  * Contact Full Name
  * Company
  * Title
  * Location
  * Company Headcount
  * Industry
  * Sales Representative
  * Company Website
* Excel Template

### For Part 2 - Upload:
* List of Sales Representatives and their Google Drive folders

### For Part 3 - Distribute:
* List of Sales Representative with the following fields:
  * Google Drive Folder
  * Email
