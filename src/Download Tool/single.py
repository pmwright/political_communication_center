import youtube_dl
import datetime
import time
import xlsxwriter

video_or_channel = input("Please enter video or channel URL: ")
xlsx_workbook_title = input("Please enter XLSX pathname: ")

ydl_opts = {}
ts = time.time()
st = datetime.datetime.fromtimestamp(ts).strftime('%m/%d/%Y %H:%M:%S')

with youtube_dl.YoutubeDL(ydl_opts) as ydl:
    each = ydl.extract_info(video_or_channel, download=False)


# meta_info = meta['entries']
wb = xlsxwriter.Workbook(xlsx_workbook_title)
ws = wb.add_worksheet('sheet1')


items = []
new_item = []
date = datetime.datetime.strptime(each['upload_date'], '%Y%m%d').strftime('%m/%d/%Y')
length = datetime.timedelta(seconds=each['duration'])

NID = each['id']
OID = ""
PCOPY_NO = ''
PCA_ID = ''
SLATE = ''
DATE = date
COMMUNICATION_TYPE = ''
PROGRAM_TYPE = ''
ELECTION_YEAR = ''
FORMAT = each['format']
POLITICAL_CONSULTANTS = each['uploader']
LENGTH = length
BEGIN_TIME = ''
FIRST_NAME = ''
LAST_NAME = ''
POLITICAL_ACTION_COMMITTEE = ''
ROLE = ''
NATION = ''
PARTY = ''
STATE = ''
OFFICE = ''
GENDER = ''
TITLE = each['title']
NOTES = ''
SUMMARY = each['description']
TRANSCRIPT = ''
SUBJECT1 = ', '.join(each['categories'])
SUBJECT2 = ''
SUBJECT3 = ''
CATDATE = st
DONOR = ''
LICENSE = each['license']
CATALOGER = ''
TAGS = ', '.join(each['tags'])

# add items into the new_item list
new_item.append(NID)
new_item.append(OID)
new_item.append(PCOPY_NO)
new_item.append(PCA_ID)
new_item.append(SLATE)
new_item.append(DATE)
new_item.append(COMMUNICATION_TYPE)
new_item.append(PROGRAM_TYPE)
new_item.append(ELECTION_YEAR)
new_item.append(FORMAT)
new_item.append(POLITICAL_CONSULTANTS)
new_item.append(LENGTH)
new_item.append(BEGIN_TIME)
new_item.append(FIRST_NAME)
new_item.append(LAST_NAME)
new_item.append(POLITICAL_ACTION_COMMITTEE)
new_item.append(ROLE)
new_item.append(NATION)
new_item.append(PARTY)
new_item.append(STATE)
new_item.append(OFFICE)
new_item.append(GENDER)
new_item.append(TITLE)
new_item.append(NOTES)
new_item.append(SUMMARY)
new_item.append(TRANSCRIPT)
new_item.append(SUBJECT1)
new_item.append(SUBJECT2)
new_item.append(SUBJECT3)
new_item.append(CATDATE)
new_item.append(DONOR)
new_item.append(LICENSE)
new_item.append(CATALOGER)
new_item.append(TAGS)

# append to the items list
items.append(new_item)

ws.write(0, 0, 'NID')
ws.write(0, 1, 'OID')
ws.write(0, 2, 'PCOPY_NO')
ws.write(0, 3, 'PCA_ID')
ws.write(0, 4, 'SLATE')
ws.write(0, 5, 'DATE')
ws.write(0, 6, 'COMMUNICATION_TYPE')
ws.write(0, 7, 'PROGRAM_TYPE')
ws.write(0, 8, 'ELECTION YEAR')
ws.write(0, 9, 'FORMAT')
ws.write(0, 10, 'POLITICAL_CONSULTANTS')
ws.write(0, 11, 'LENGTH')
ws.write(0, 12, 'BEGIN_TIME')
ws.write(0, 13, 'FIRST_NAME')
ws.write(0, 14, 'LAST_NAME')
ws.write(0, 15, 'POLITICAL_ACTION_COMMITTEE')
ws.write(0, 16, 'ROLE')
ws.write(0, 17, 'NATION')
ws.write(0, 18, 'PARTY')
ws.write(0, 19, 'STATE')
ws.write(0, 20, 'OFICE')
ws.write(0, 21, 'GENDER')
ws.write(0, 22, 'TITLE')
ws.write(0, 23, 'NOTES')
ws.write(0, 24, 'SUMMARY')
ws.write(0, 25, 'TRANSCRIPT')
ws.write(0, 26, 'SUBJECT1')
ws.write(0, 27, 'SUBJECT2')
ws.write(0, 28, 'SUBJECT3')
ws.write(0, 29, 'CATDATE')
ws.write(0, 30, 'DONOR')
ws.write(0, 31, 'LICENSE')
ws.write(0, 32, 'CATALOGER')
ws.write(0, 33, 'TAGS')


i = 0
for row in items:
    j = -1
    i += 1
    for col in row:
        j += 1
        ws.write(i, j, str(col))

print('Excel writing is done!!!!')
wb.close()