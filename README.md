# Distribute data using python

* Split and distribute **contacts** to **sales representative** automatically

## Tools:
1. Python
2. Google API (GDrive, GSheet)

# Notes:
* 1 GSheet = Owner

## Code Logic

1. Create file by Teamlink
4. Upload to GDrive - Python, GDrive API
5. Freeze top visible row and column - Python, GSheet API


## Email Code logic


1. Attachment
2. Log sent messages


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
