# genThemeSheet.py
# generates xlsx file with all the songs in PCO db sorted by theme for use in Pastor's binder.
# in: nothing
# out: themeSheet.xlsx

import requests
import json
import re
import csv
import pandas as pd       # for V1 (exports csv)
import xlsxwriter           # for V2 (exports formatted xlsx)
import time
from tqdm import tqdm

##############################################
# REPLACE THIS WITH YOUR PCO API AUTHORIZATION
##############################################
head = {"Authorization":"Basic MmVhYTc2NTVkYzExZDFjNzFhODI5NmQ2ODkyMmE0MTAwOTkxZDQ2NmNjYzM1ZmJhOWZjOGMxZWQyZDI5MWUxZjphMmZhZWEyNDc1MmZhMTRjYzEzM2UzNjRlMmFjM2IzMzEyYmI5OWEzYWY0ZTEyNTEzODg2NzJjZTQ4ZTNlZmYy"}
# head = {"Authorization":"Your_API_Key_HERE"}

allTags = {}
allSongs = {}


def getSongInfo(link):
    songReq = requests.get('{0}/arrangements'.format(link), headers=head)
    songObj = songReq.json()
    if 'errors' in songObj:
        time.sleep(20)
        songReq = requests.get('{0}/arrangements'.format(link), headers=head)
        songObj = songReq.json()

    return songObj


def getFirstLine(songObj):
    try:
        firstLine = songObj["data"][0]["attributes"]["lyrics"]
        firstLine = firstLine.splitlines()[1]
        firstLine = re.sub('(\\[.*?\\])', '', firstLine)
    except:
        firstLine = ""
    return firstLine


def getBPM(songObj):
    BPM = songObj["data"][0]["attributes"]["bpm"]
    return BPM


def getKeys(songObj):
    keyList = []
    for arr in songObj["data"]:
        key = arr["attributes"]["chord_chart_key"]
        keyList.append(key)
    return keyList


def getTags(link):
    tagsReq = requests.get('{0}/tags'.format(link), headers=head)
    tagsObj = tagsReq.json()
    if 'errors' in tagsObj:
        time.sleep(20)
        tagsReq = requests.get('{0}/tags'.format(link), headers=head)
        tagsObj = tagsReq.json()

    return tagsObj


def updateAllTags(tagsObj, title):
    for songTag in tagsObj["data"]:
        tag = songTag["attributes"]["name"]

        if tag not in allTags:
            allTags[tag] = []

        allTags[tag].append(title)


def sortTags():
    for tag in allTags:
        allTags[tag].sort()


# For each song in that set... ###########################################################################
def getSongData(song):
    title = song["attributes"]["title"]
    link = song["links"]["self"]

    allSongs[title] = {}

    songObj = getSongInfo(link)

    allSongs[title]["firstLine"] = getFirstLine(songObj)

    allSongs[title]["BPM"] = getBPM(songObj)

    allSongs[title]["keys"] = getKeys(songObj)

    tagsObj = getTags(link)

    updateAllTags(tagsObj, title)
######################################################################################################


# Update the theme sheet (filename) with new songs
def updateThemeSheet(filename):
    return


# Create a brand new theme sheet directly from PCO db
def generateNewThemeSheet():
    # Fetch songs
    offset = 0
    page = 100
    while True:
        songsReq = requests.get('https://api.planningcenteronline.com/services/v2/songs/?order=title&where[hidden]=false&per_page={}&offset={}'.format(page, offset), headers=head)
        songsObj = songsReq.json()

        if len(songsObj["data"]) == 0:
            break

        print 'Generating for offset {} with {} per page...'.format(offset, page)
        for song in tqdm(songsObj["data"]):
            getSongData(song)

        offset += page


##########################################################################################################
# Version 1 generates a csv file using pandas
def genCSV():
    df = pd.DataFrame()

    for tag in allTags:
        # apply formatting for section header
        df = df.append([tag], ignore_index=True)

        for song in allTags[tag]:
            # apply formatting for song
            df = df.append([[ song, allSongs[song]["firstLine"], allSongs[song]["BPM"], allSongs[song]["keys"] ]], ignore_index=True)

    df.to_csv(open('themeSheet.csv', 'wb'), sep='\t', encoding='utf-8')
##########################################################################################################


##########################################################################################################
# Version 2 generates xls file using xlsxwriter
def genXLS():

    workbook = xlsxwriter.Workbook('themeSheet.xlsx')
    worksheets = {}

    # Style/formatting stuff
    sectionStyle = workbook.add_format({'font_name': 'Avenir', 'bg_color': 'black', 'font_color': 'white', 'font_size': 10})
    songCenterStyle = workbook.add_format({'font_name': 'Avenir', 'bottom': 4, 'align': 'center', 'font_size': 10})
    songStyle = workbook.add_format({'font_name': 'Avenir', 'bottom': 4, 'font_size': 10})

    # getting tag groups
    # allTagGroups = { idOfTagGroup --> alphabetical list of ids of children tags }
    tagGroups = requests.get('https://api.planningcenteronline.com/services/v2/tag_groups', headers=head)
    tagGroupObj = tagGroups.json()
    allTagGroups = {}
    for tagGroup in tagGroupObj["data"]:
        if tagGroup["attributes"]["tags_for"] == 'song':
            worksheets[tagGroup["attributes"]["name"]] = workbook.add_worksheet()

            allTagGroups[tagGroup["attributes"]["name"]] = []
            r = requests.get('https://api.planningcenteronline.com/services/v2/tag_groups/{0}/tags'.format(tagGroup["id"]), headers=head)
            groupData = r.json()
            for tag in groupData["data"]:
                allTagGroups[tagGroup["attributes"]["name"]].append(tag["attributes"]["name"])

            sorted(allTagGroups[tagGroup["attributes"]["name"]])

    for tagGroup in allTagGroups:
        worksheet = worksheets[tagGroup]    # creates new sheet for each tag group (genre, theme, purpose, mood)
        worksheet.hide_gridlines(2)
        worksheet.set_column('A:A', .55)
        worksheet.set_column('B:B', 34)
        worksheet.set_column('C:C', 46)
        worksheet.set_column('D:D', 5)
        worksheet.set_column('E:E', 5)
        worksheet.set_column('F:F', .18)
        worksheet.set_default_row(16)

        row = 0
        for tag in allTagGroups[tagGroup]:
            worksheet.merge_range(row, 0, row, 4, tag, sectionStyle)
            row += 1

            try:
                for song in allTags[tag]:
                    if row % 50 == 0:
                        worksheet.merge_range(row, 0, row, 4, tag + ' (cont.)', sectionStyle)
                        row += 1

                    worksheet.write(row, 1, song, songStyle)
                    worksheet.write(row, 2, allSongs[song]["firstLine"], songStyle)
                    worksheet.write(row, 3, allSongs[song]["BPM"], songCenterStyle)
                    # regex gets rid of unicode and random quotes and stuff
                    keys = re.sub('([\\[\\]\'u])', '', str(allSongs[song]["keys"]))
                    worksheet.write(row, 4, keys, songCenterStyle)
                    row += 1
            except:
                continue
    workbook.close()

##########################################################################################################
def main():
    mode = raw_input("Type \'update\' to update theme sheet. Type \'new\' to generate a brand new theme sheet. (You will have to do a lot of manual work if you make a new one)\n")
    outputType = raw_input("Type \'xls\' to generate a formatted xlsx file. Type \'csv\' to generate an ugly csv file.\n")
    if (mode == 'update'):
        filename = raw_input("Please type the filename of the old theme sheet\n")
        updateThemeSheet(filename)
    elif (mode == 'new'):
        print "Generating new theme sheet... "
        generateNewThemeSheet()
    else:
        print 'Not a valid option'
        exit(1)


    if outputType == 'csv':
        genCSV()
    else:
        genXLS()

main()
