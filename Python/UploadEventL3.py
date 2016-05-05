# -*- coding: utf-8 -*-

import openpyxl
import sys
import codecs
import re

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime
import time

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'ScriptsTimeTable'

CalID = {
    'L1' : {
        'Tous' : '9fnaccpmdsbu40ikv0t05oiui0@group.calendar.google.com',
        'G1' : 'd5p6tiig01oa6eeavmmjo8hrv4@group.calendar.google.com',
        'G2' : 'v66271c0mdgufkchar1ddal66k@group.calendar.google.com',
    },
    'L2' : {
        'Tous' : '50isphb3vtjlltfk5an9sdi05k@group.calendar.google.com',
        'G1' : '4tftm0b1i9lvtq90hsnku80fro@group.calendar.google.com',
        'G2' : 'dajefooh8hn892un0a5mfafeag@group.calendar.google.com',
    },
    'L3' : {
        'BioIng' : {
            'G1' : 'dj40eu0osq3tt95s41fd11q0ig@group.calendar.google.com',
            'G2' : 'ski1p038ldoee5vf0srvv77inc@group.calendar.google.com',
            'Tous' : 'q3gsg7ano4njt9rr6puo2p5cus@group.calendar.google.com'
        },
        'Commun': {
            'G1' : '7ib1d8a2utbvjbk1idv2vjhit4@group.calendar.google.com',
            'G2' : '7ou0h6ao1rgip0ciri51gpoa7k@group.calendar.google.com',
            'Tous' : 'gcs5fsuuiqesp9cf840hbtmtp0@group.calendar.google.com',
        },
        'IPMI' : {
            'G1' : 'f4vrim0m300fuh3hfg3otj30r4@group.calendar.google.com',
            'G2' : '4f0uetqe8kjj00nj0u68k7mo94@group.calendar.google.com',
            'Tous' : 't141mkpui8u1eff7t7rh4fg064@group.calendar.google.com',
        },
    },
}

ThisYear = 'L3'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials

sys.stdout = codecs.getwriter('utf8')(sys.stdout)
sys.stderr = codecs.getwriter('utf8')(sys.stderr)

def iter_range(rows, sortedIdxLst):
    try:
        minIdx = sortedIdxLst.next()
    except StopIteration:
        return
    idx = 0
    for row in rows:
        idx += 1
        if idx < minIdx:
            continue
        yield row
        try:
            minIdx = sortedIdxLst.next()
        except StopIteration:
            return

def iter_from(i):
    while True:
        yield i
        i += 1

def load_event(filename, IdFile):
    wb = openpyxl.load_workbook(filename=filename, read_only=True)

    ws = wb['Listing']

    UnbiasedWHL = [u'Date', u'Début', u'Fin', u'Libellé', u'Salle', u'Enseignant', u'UE', u'Discipline', u'Groupes', u'Commentaires']
    WantedHeadLst = []
    HeadRange = []
    Remember = []
    for row in iter_range(ws.rows, iter([1])):
        for thisCell in row:
            if thisCell.value in UnbiasedWHL:
                HeadRange += [openpyxl.cell.column_index_from_string(thisCell.column)]
                WantedHeadLst += [thisCell.value]

    endline = re.compile(r'_x0*D_')

    locationre = re.compile(r'(.*?) # (.*?)\n(.*?)\n(.*)(\n.*|)')

    for row in iter_range(ws.rows, iter_from(2)):
        idxWHL = 0
        event = {
            'recurrence': [],
            'attendees': [],
            'reminders': {
                'useDefault': True,
                'overrides': [],
            },
        }
        seDate, seDebut, seFin = "", "", ""
        summary, rooms, teacher = "", "", ""
        UE, discipline, groups = "", "", ""
        comments, teachers = "", ""
        for cell in iter_range(row, iter(HeadRange)):
            if cell.value != None and cell.value != "":
                t = unicode(cell.value)
                u = re.sub(endline, '', t)
                if WantedHeadLst[idxWHL] == UnbiasedWHL[0]:
                    seDate = u[:10]
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[1]:
                    seDebut = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[2]:
                    seFin = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[3]:
                    summary = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[4]:
                    rooms = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[5]:
                    teacher = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[6]:
                    UE = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[7]:
                    discipline = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[8]:
                    groups = u
                elif WantedHeadLst[idxWHL] == UnbiasedWHL[9]:
                    comments = u
                else:
                    print "notfound = True"
            idxWHL += 1
        event['location'] = rooms
        if summary == '':
            continue
        lm = locationre.match(summary)
        teachers = ''
        if lm != None:
            teachers = lm.group(4)
            event['description'] = teachers
            if comments != "":
                event['description'] += ' ' + comments
        else:
            teachers = ""
            event['description'] = ""
        allsummary = []
        for k in [UE, rooms, discipline, teachers, comments]:
            if k != '':
                allsummary += [k]
        event['summary'] = " ".join(allsummary)
        event['libl'] = summary
        if groups == u'1/2':
            ThisCal = 'G1'
        elif groups == u'2/2':
            ThisCal = 'G2'
        else:
            ThisCal = 'Tous'

        event['start'] = {'dateTime' : seDate + 'T' + seDebut, 'timeZone': 'Europe/Paris',}
        event['end'] = {'dateTime' : seDate + 'T' + seFin, 'timeZone': 'Europe/Paris',}
        print "CALID", ThisYear, ThisCal, groups
        print event
        Remember += [ [event, ThisCal, IdFile] ]
    return Remember

def aaaaaa(event, ThisCal, ThisSuperCal):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    print event, ThisCal, ThisSuperCal
    eventsResult = service.events().list(
        calendarId=CalID[ThisYear][ThisSuperCal][ThisCal],
        timeMin=event['start']['dateTime']+'+02:00',
        timeMax=event['end']['dateTime']+'+02:00',
        maxResults=10,
        singleEvents=True,
        orderBy='startTime').execute()
    eventsRetrieved = eventsResult.get('items', [])
    foundEvent=False
    print len(eventsRetrieved)
    if len(eventsRetrieved) > 0:
        for eee in eventsRetrieved:
            if 'summary' in eee.keys() and (eee['summary'] == event['summary'] or unicode(eee['summary']).replace("\n", " ") == unicode(event['summary']).replace("\n", " ")) and 'creator' in eee.keys() and eee['creator']['email'] == u'scripts.villebon@gmail.com':
                if ('description' in eee.keys() and eee['description'] != event['description'] and unicode(eee['description']).replace("\n", " ") != unicode(event['description']).replace("\n", " ") ) or ('summary' in eee.keys() and eee['summary'] != event['summary'] and eee['summary'].replace('\n', ' ') != event['summary'].replace("\n", " ")):
                    if not groups in [ '2/3', '3/3', '2/4', '3/4', '4/4' ]:
                        print "PATCH", eee
                        eee['description'] = event['description']
                        eee['summary'] = event['summary']
                        print "INTO", eee, "in", ThisYear, ThisSuperCal, ThisCal
                        pevent = service.events().patch(calendarId=CalID[ThisYear][ThisSuperCal][ThisCal], eventId=eee['id'], fields="description,summary", body=eee).execute()
                        foundEvent=True
                    else:
                        print "DELETE", eee, "in", ThisYear, ThisSuperCal, ThisCal
                        pevent = service.events().delete(calendarId=CalID[ThisYear][ThisSuperCal][ThisCal], eventId=eee['id']).execute()
                    break
                else:
                    print "FOUND", "in", ThisYear, ThisSuperCal, ThisCal
                    foundEvent = True
            else:
                print "NOT IN PREVIOUS", eee
    if not foundEvent:
        print "CREATE", event, "in", ThisYear, ThisSuperCal, ThisCal
        oevent = service.events().insert(calendarId=CalID[ThisYear][ThisSuperCal][ThisCal], body=event).execute()


R = []
R += load_event('L3BIOING_v3.xlsm', 'BioIng')
R += load_event('L3IPMI_v3.xlsm', 'IPMI')

R2 = sorted(R, key = lambda x: time.strftime("%s", time.strptime(x[0]['start']['dateTime'], "%Y-%m-%dT%H:%M:%S")) + time.strftime("%s", time.strptime(x[0]['end']['dateTime'], "%Y-%m-%dT%H:%M:%S")) + x[1] + x[2])

R3 = []
passNext=False
for i, r in enumerate(R2):
    if passNext:
        passNext=False
        continue
    elif i == len(R2) - 1:
        print "Last"
        aaaaaa(r[0], r[1], r[2])
        #R3 += [r]
    elif r[0] == R2[i+1][0] and r[1] == R2[i+1][1] and r[2] != R2[i+1][2]:
        print "Commun"
        aaaaaa(r[0], r[1], 'Commun')
        #R3 += [[r[0], r[1], 'Commun']]
        passNext=True
    else:
        print "Not common because", r, "!=", R2[i+1]
        if (r[1] != 'G2' and r[1] != 'G1') or r[2] != 'BioIng':
            aaaaaa(r[0], r[1], r[2])
        else:
            print "ERROR"
            aaaaaa(r[0], r[1], r[2])
        #R3 += [r]


