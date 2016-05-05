# -*- coding: utf-8 -*-

from openpyxl import *
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

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'ScriptsTimeTable'

CalID = {
    'L1' : {
        'Tous' : '***@group.calendar.google.com',
        'G1' : '***@group.calendar.google.com',
        'G2' : '***@group.calendar.google.com',
    },
    'L2' : {
        'Tous' : '***@group.calendar.google.com',
        'G1' : '***@group.calendar.google.com',
        'G2' : '***@group.calendar.google.com',
    },
    'L3' : {
        'Biolng' : '***@group.calendar.google.com',
        'G1' : '***@group.calendar.google.com',
        'G2' : '***@group.calendar.google.com',
        'IPM1G1' : '***@group.calendar.google.com',
        'IPMIG2' : '***@group.calendar.google.com',
        'IPMITous' : '***@group.calendar.google.com',
        'Tous' : '***@group.calendar.google.com'
    },
}

ThisYear = 'L2'

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
    

wb = load_workbook(filename='L2P17_v4.xlsm', read_only=True)

ws = wb['Listing']

UnbiasedWHL = [u'Date', u'Début', u'Fin', u'Libellé', u'Salle', u'Enseignant', u'UE', u'Discipline', u'Groupes', u'Commentaires']
WantedHeadLst = []
HeadRange = []

for row in iter_range(ws.rows, iter([1])):
    for thisCell in row:
        if thisCell.value in UnbiasedWHL:
            HeadRange += [cell.column_index_from_string(thisCell.column)]
            WantedHeadLst += [thisCell.value]

endline = re.compile(r'_x0*D_')

locationre = re.compile(r'(.*?) # (.*?)\n(.*?)\n(.*)(\n.*|)')

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
service = discovery.build('calendar', 'v3', http=http)


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
        if cell.value != None:
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
                pass
        idxWHL += 1
    event['location'] = rooms
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

    eventsResult = service.events().list(
        calendarId=CalID[ThisYear][ThisCal],
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
                        print "INTO", eee
                        pevent = service.events().patch(calendarId=CalID[ThisYear][ThisCal], eventId=eee['id'], fields="description,summary", body=eee).execute()
                        foundEvent=True
                    else:
                        print "DELETE", eee
                        pevent = service.events().delete(calendarId=CalID[ThisYear][ThisCal], eventId=eee['id']).execute()
                    break
                else:
                    print "FOUND"
                    foundEvent = True
            else:
                print "NOT IN PREVIOUS", eee
    if not foundEvent:
        print "CREATE", event
        oevent = service.events().insert(calendarId=CalID[ThisYear][ThisCal], body=event).execute()


