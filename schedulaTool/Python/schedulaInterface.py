# schedulaInterface.py implments a python interface to the admin appointments functionallity of https://schedula.sportstg.com/

# MIT License
#
# Copyright (c) 2020 Ian Crossing
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


# imports
import requests
import urllib
import sys
import csv
import re
import calendar
from datetime import date
from datetime import timedelta
import time
import traceback

# TODO: parallelise requests to speedup

################################################################################
#########################           FUNCTIONS          #########################
################################################################################
# getPage returns the requested page from schedula
#
# session       - session to use for connection to schedula
# url           - URL to get
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# returns the output of session.get(url)
def getPage(session, url, proxy=False, proxyDict={}):
    # TODO: check session is still loged in
    if proxy:
        return session.get(url, proxies=proxyDict, verify=False)
    else:
        return session.get(url)

# post performs an HTTP post request
#
# session       - session to use for connection to schedula
# url           - URL to get
# data          - data segment of POST request
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# returns the output of session.post(url)
def post(session, url, data, proxy=False, proxyDict={}):
    # TODO: check session is still loged in
    headers = { "content-type" : "application/x-www-form-urlencoded", "Accept-Language" : "en-US,en;q=0.5", "Origin" : "https://schedula.sportstg.com"}
    if proxy:
        r = session.post(url, data=data, headers=headers, proxies=proxyDict, verify=False)
    else:
        r = session.post(url, data=data, headers=headers)

    # debug
    referees = getRefsFromText(r.text)

    return r

# getSession returns an html session to be used for connection to schedula
#
# user          - username
# pswd          - password
# proxy         - use an http proxy if true
# proxyAddress  - address of proxy
#
# Returns a session logged in to schedula, throws error on fail
def getSession(user, pswd, proxy=False, proxyDict={}):
    print('Attempting login...')
    
    # Session object
    session = requests.Session()

    # url of schedula login page
    url = 'https://schedula.sportstg.com'

    # get login page
    r = getPage(session, url, proxy, proxyDict)

    # login data
    data = 'xjxfun=dologin&xjxargs[]=' + urllib.parse.quote('<xjxobj><e><k>email</k><v>S<![CDATA[' + user + ']]></v></e><e><k>password</k><v>S' + pswd + '</v></e><e><k>btnlogin</k><v>SLogin</v></e></xjxobj>')

    # attempt login
    r2 = post(session, url, data, proxy, proxyDict)

    # check login status
    text = r2.text
    text = text.split('CDATA')
    text = text[1].split('"')
    text = text[1]

    if text == 'https://schedula.sportstg.com/index.php?action=dashboard':
        print('Login Success')
    else:
        print('Login Fail')
        raise Exception('Login Failed. Incorrect username or password')

    return session

# getOrganisations returns a list of organisations and ids
#
# session       - session to use for connection to schedula
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# Returns a list of organisation and id pairs
def getOrganisations(session, proxy=False, proxyDict={}):
    # get an admin page
    url = 'https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_by_week'
    r3 = getPage(session, url, proxy, proxyDict)

    # the interesting bit is below. This gives the organisations, their id and which function to use to get the seasons
    # <form name="search_fixture">
    # ...
    # <td><div ... name="orgs" id="orgs" class="input_select" 
    #      onchange="xajax_GetSeasons(document.search_fixture.orgs.value,'AppointByWeek')">
    #  <option value=""></option>
    #  <option value="664">AHJSA</option>
    #  <option value="208">FFACUP</option>
    #  <option value="9">FFSA</option>
    #  <option value="97">NPLSA</option>
    # </div></td>
    
    # parse the html into useful tokens
    text = r3.text
    text = text.split('<form name="search_fixture">')
    text = text[1]
    text = text.split('<')

    # extract organisation values and strings
    organisations = []
    for i in range(len(text)):
        # check if text is a competition option
        if text[i].startswith('option'):
            try:
                t = text[i]
                orgNum = t.split('"')[1]
                orgName = (t.split('>')[1]).split('<')[0]
                org = [orgNum, orgName]            
                # Remove blank option
                if orgNum != "":
                    organisations.append(org)
                    #print('  Organisation found: ' + orgNum + ' - ' + orgName)
            except: # ignore blank ones
                pass
        
        # check for the end of the form
        elif text[i].startswith('/select'):
            break

    return organisations

# getSeasons returns the season and season id's for the given organisation
#
# session       - session to use for connection to schedula
# orgId         - Id of the organisation to query
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# Returns a list of season, id pairs
def getSeasons(session, orgId, proxy=False, proxyDict={}):
    # HTTP data
    url = 'https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_by_week'
    data = 'xjxfun=GetSeasons&xjxargs[]=S' + orgId + '&xjxargs[]=SAppointByWeek'
    r4 = post(session, url, data, proxy, proxyDict)

    # example response below. Note we get the season names and id + the function to get the season weeks
    # <?xml version="1.0" encoding="utf-8" ?><xjx> ...
    # <select ... name="season" id="season" ... onchange="xajax_GetSeasonWeeks(document.search_fixture.season.value,'AppointByWeek')">
    # <option value=""></option>
    # <option value="3033">2020</option>
    # <option value="2817">2019</option>
    # ]]></cmd></xjx>

    # split into useful tokens
    text = r4.text
    text = text.split('</option>')

    seasons = []
    for i in range(len(text)):
        if i==0: # ignore the first one
            pass
        else:
            if ']]' in text[i]: # check for end of list
                break
            seasonID = text[i].split('"')[1]
            seasonName = text[i].split('>')[1]
            seasons.append([seasonID, seasonName])

    return seasons


# getSeasonWeeks returns the weeks in the given season
#
# session       - session to use for connection to schedula
# seasonID      - Id of the season to query
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# Returns a list of week, id pairs
def getSeasonWeeks(session, seasonID, proxy=False, proxyDict={}):
    # HTTP data
    url = 'https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_by_week'
    data = 'xjxfun=GetSeasonWeeks&xjxargs[]=S' + seasonID + '&xjxargs[]=SAppointByWeek'
    r5 = post(session, url, data, proxy, proxyDict)

    # example response below. Note we are told the week strings, ids and the function to get the fixtures
    # <?xml version="1.0" encoding="utf-8" ?><xjx>...
    # onclick="xajax_ShowFixturesForWeek(3070,document.search_fixture.week.value)"...
    # id="week" name="week">
    # <option value="2020-01-13_2020-01-19">Week 1 (Jan 13 to Jan 19)</option>
    # <option value="2020-01-20_2020-01-26">Week 2 (Jan 20 to Jan 26)</option>
    # ...
    # <option value="2020-10-26_2020-11-01">Week 34 (Oct 26 to Nov 1)</option>
    # </select>]]></cmd></xjx>
    
    # split into tokens
    text = r5.text.split('<')

    # get the weeks
    weeks = []
    for i in range(len(text)):
        if text[i].startswith('option'): # check for a week in the token
            try:
                t = text[i]
                weekID = t.split('"')[1]
                weekName = (t.split('>')[1])
                weeks.append([weekID, weekName])
            except:
                pass
        elif text[i].startswith('/select'): # no more weeks
            break

    return weeks


# getSeasonWeeks returns the fixtures for the given week in the given season
#
# session       - session to use for connection to schedula
# seasonID      - Id of the season to query
# weekID        - Id of the week to query
# proxy         - use an http proxy if true
# proxyDict     - address of proxy
#
# Returns a list of fixture details [compName, date, time, home, away, ground, fixID]
def getFixturesForWeek(session, seasonID, weekID, proxy=False, proxyDict={}):
    # HTTP data
    data = 'xjxfun=ShowFixturesForWeek&xjxargs[]=N' + seasonID + '&xjxargs[]=S' + weekID
    url = 'https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_by_week'
    r6 = post(session, url, data, proxy, proxyDict)

    # response is a big table with each competition + some other random stuff

    # split into tokens
    text = r6.text.split('<')

    # parse the table
    fixtures = []
    fixture = ['compName', 'date', 'time', 'home', 'away', 'ground', 'fixID']
    for i in range(len(text)):
        row = text[i]

        # check for end of table
        if row.startswith('/table>'):
            break

        # check for new competition
        elif row.startswith('th colspan='):
            #print("New comp")
            fixture[0] = row.split('>')[1]
            fixture[1] = ''
            fixture[2] = ''
            fixture[3] = ''
            fixture[4] = ''
            fixture[5] = ''
            fixture[6] = ''

        # Check for data
        elif row.startswith('td'):
            #print("Insert data")
            try:
                rowData = row.split('>')[1]

                if rowData != 'v':
                    if rowData != '&nbsp;':
                        if fixture[1] == '':
                            fixture[1] = rowData
                        elif fixture[2] == '':
                            fixture[2] = rowData
                        elif fixture[3] == '':
                            fixture[3] = rowData
                        elif fixture[4] == '':
                            fixture[4] = rowData
                        elif fixture[5] == '':
                            fixture[5] = rowData
            except:
                pass

        # Check for fixture ID
        elif row.startswith('u style='):
            #print('Fix ID')
            fixture[6] = (row.split('fixtureid=')[1]).split('&')[0]

            # Add the fixture
            fixtures.append(list(fixture))

            # clear the fixture
            fixture[1] = ''
            fixture[2] = ''
            fixture[3] = ''
            fixture[4] = ''
            fixture[5] = ''
            fixture[6] = ''

    # TODO: filter out U8-U11 to save time

    return fixtures


# pullAll gets all the fixtures and appointments for the given year and outputs two csv's
#
# session            - session to use for connection to schedula
# year               - calander year to pull from, e.g. '2020'
# fixturesFile       - file name to store fixtures in
# appointmentsFile   - file name to store appointments in
# useProxy           - use an http proxy if true
# proxyDict          - address of proxy
#
# returns - {'fixturesList':fixtures,'appointmentList':appointments}
#       fixtures     - ['OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date','Time','Home','Away','Ground','FixtureID']
#       appointments - ['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']
def pullAll(session, year='2020', fixturesFile='Fixtures.csv', appointmentsFile='Appointments.csv', useProxy=False, proxyDict={}):
    ##################
    ##   Fixtures   ##
    ##################
    # Get Organisations and ID's
    # Getting organisations
    print('Getting organisations...')

    organisations = getOrganisations(session, useProxy, proxyDict)

    print('Found ' + str(len(organisations)) + ' Organisations')
    print(organisations)

    # Step 1: Get seasons and ID's
    print('Getting Seasons...')

    seasons = []

    for org in organisations:
        orgNum = org[0]
        orgName = org[1]
        s = getSeasons(session, orgNum, useProxy, proxyDict)

        for se in s:
            seasons.append([orgNum, orgName, se[0], se[1]])

    print('Found Seasons:')
    print("['OrgID','Org','SID','SName']")
    print(seasons)

    # select the requested year
    if year != '':
        removes = []
        for s in seasons:
            if s[3] != year:
                removes.append(s)

        for r in removes:
            seasons.remove(r)


    # Step 2: Get season weeks (based on season ID)
    print('Getting season weeks...')
    weeks = []
    for season in seasons:
        seasonID = season[2]
        ws = getSeasonWeeks(session, seasonID, useProxy, proxyDict)

        for w in ws:
            weeks.append([season[0], season[1], season[2], season[3], w[0], w[1]])



    print('Found Weeks:')
    print("['OrgID','Org','SID','SName','WID','WName']")
    for week in weeks:
        print(week)

    # Step 3: Get Fixtures for week (based on season and week)
    print("\nGetting Fixtures...")
    fixture = ['OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date','Time','Home','Away','Ground','FixtureID']
    fixtures = []
    #fixtures.append(list(fixture))
    perCent = 1 / float(len(weeks))
    count = 0

    for week in weeks:

        orgNum = week[0]
        orgName = week[1]
        seasonID = week[2]
        seasonName = week[3]
        weekID = week[4]
        weekName = week[5]

        # get the fixtures for the week
        fixs = getFixturesForWeek(session, seasonID, weekID, useProxy, proxyDict)

        for f in fixs:
            # check balcklist
            if not isBlacklisted(f[0]):
                fixtures.append([orgNum,orgName,seasonID,seasonName,weekID,weekName,f[0],f[1],f[2],f[3],f[4],f[5],f[6]])

        count = count+1
        total = count * perCent
        print('\r',end='')
        print("%.2f" % (total*100), end='')
        print('%',end='')


    print('\n')
    print("Found " + str(len(fixtures)) + " Fixtures\n")

    # output fixtures to CSV
    if fixturesFile != '':
        output = []
        output.append(fixture)
        for f in fixtures:
            output.append(list(f))
        with open(fixturesFile,"w",newline='') as file:
            writer = csv.writer(file)
            writer.writerows(output)
        print ("Output fixtures to CSV")

    ##################
    ## Appointments ##
    ##################
    if(len(fixtures) == 0):
        return

    print("Getting appointments...\n")
    appointments = []
    appointments.append(['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus'])
    perCent = 1/float(len(fixtures))
    count = 0
    for f in fixtures:
        fixtureID = f[12]
        match = lookupFixture(session, fixtureID)

        if len(match) != 0:
            appointments.extend(match.copy())
        count = count+1
        total = count * perCent

        print("\r",end='')
        print("%.2f" % (total*100), end='')
        print("%", end='')

    print("Found " + str(len(appointments)-1) + " appointments")

    # save the appointments to a csv
    if appointmentsFile != '':
        with open(appointmentsFile,"w",newline='') as file:
            writer = csv.writer(file)
            writer. writerows(appointments)
        print("Output appointments to csv")

    if len(appointments) > 1:
        return {'fixturesList':fixtures,'appointmentList':appointments[1:]}
    else:
        return {'fixturesList':fixtures,'appointmentList':[]}

# update28 gets all the fixtures and appointments for the next 28 days and updates or creates the csv
#
# session            - session to use for connection to schedula
# year               - calander year to pull from, e.g. '2020'
# fixturesFile       - file name to store fixtures in
# appointmentsFile   - file name to store appointments in
# useProxy           - use an http proxy if true
# proxyDict          - address of proxy
#
# returns - {'fixturesList':fixtures,'appointmentList':appointments}
#       fixtures     - ['OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date','Time','Home','Away','Ground','FixtureID']
#       appointments - ['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']
def update28(session, year='', fixturesFile='Fixtures28.csv', appointmentsFile='Appointments28.csv', startDate=date.today(), numberDays=28, useProxy=False, proxyDict={}):
    ##################
    ##   Fixtures   ##
    ##################
    # Get Organisations and ID's
    # Getting organisations
    print('Getting organisations...')

    organisations = getOrganisations(session, useProxy, proxyDict)

    print('Found ' + str(len(organisations)) + ' Organisations')
    print(organisations)

    # Step 1: Get seasons and ID's
    print('Getting Seasons...')

    seasons = []

    for org in organisations:
        orgNum = org[0]
        orgName = org[1]
        s = getSeasons(session, orgNum, useProxy, proxyDict)

        for se in s:
            seasons.append([orgNum, orgName, se[0], se[1]])

    print('Found Seasons:')
    print("['OrgID','Org','SID','SName']")
    print(seasons)

    # select the requested year
    if year != '':
        removes = []
        for s in seasons:
            if s[3] != year:
                removes.append(s)

        for r in removes:
            seasons.remove(r)


    # Step 2: Get season weeks (based on season ID)
    print('Getting season weeks...')
    # Calculate end date - use 30 days to ensure complete coverage
    startDate = startDate - timedelta(days=1)
    endDate = startDate + timedelta(days=(numberDays+2))

    #print('numDays = ' + str(numberDays))
    #print('s: ' + str(startDate) + ' e: ' + str(endDate))

    weeks = []
    for season in seasons:
        seasonID = season[2]
        ws = getSeasonWeeks(session, seasonID, useProxy, proxyDict)

        for w in ws:
            weekID = w[0]
            dateStrings = weekID.split('_')
            # grab the week either side of the date range, just to make sure
            wstart = stringToDate(dateStrings[0])
            wstart2 = stringToDate(dateStrings[0]) - timedelta(days=7)
            wend = stringToDate(dateStrings[1])
            wend2 = stringToDate(dateStrings[1]) + timedelta(days=7)

            if ((wstart >= startDate) and (wstart <= endDate)) or ((wend >= startDate) and (wend <= endDate)) or ((wstart2 >= startDate) and (wstart2 <= endDate)) or ((wend2 >= startDate) and (wend2 <= endDate)):
                weeks.append([season[0], season[1], season[2], season[3], w[0], w[1]])


    print('Looking at Weeks:')
    print("['OrgID','Org','SID','SName','WID','WName']")
    for week in weeks:
        print(week)

    # Step 3: Get Fixtures for week (based on season and week)
    print("\nGetting Fixtures...")
    fixture = ['OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date','Time','Home','Away','Ground','FixtureID']
    fixtures = []
    reject = []
    #fixtures.append(list(fixture))
    perCent = 1 / float(len(weeks))
    count = 0
    mounthToInt = dict((v,k) for k,v in enumerate(calendar.month_abbr))
    for week in weeks:

        orgNum = week[0]
        orgName = week[1]
        seasonID = week[2]
        seasonName = week[3]
        weekID = week[4]
        weekName = week[5]

        # get the fixtures for the week
        fixs = getFixturesForWeek(session, seasonID, weekID, useProxy, proxyDict)

        for f in fixs:
            # check balcklist
            if not isBlacklisted(f[0]):
                # get fixture date
                fixtureDateStrs = f[1].split(' ') # e.g. ['Sat', 'Feb', '8']
                day = int(fixtureDateStrs[2])
                month = int(mounthToInt[fixtureDateStrs[1]])
                year = int(seasonName)
                fixtureDate = date(year,month,day)

                if (fixtureDate >= startDate and fixtureDate <= endDate):
                    fixtures.append([orgNum,orgName,seasonID,seasonName,weekID,weekName,f[0],f[1],f[2],f[3],f[4],f[5],f[6]])
            else:
                reject.append(f[0])

        count = count+1
        total = count * perCent
        print('\r',end='')
        print("%.2f" % (total*100), end='')
        print('%',end='')


    print('\n')
    print("Found " + str(len(fixtures)) + " Fixtures after rejecting " + str(len(reject)) + ". Competitions rejected:")

    # get unique and print
    printList = []
    for r in reject:
        if r not in printList:
            printList.append(r)
    print(printList)


    ##################
    ## Appointments ##
    ##################
    if(len(fixtures) == 0):
        return

    print("\nGetting appointments...\n")
    appointments = []
    appointments.append(['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus'])
    perCent = 1/float(len(fixtures))
    count = 0
    for f in fixtures:
        fixtureID = f[12]
        match = lookupFixture(session, fixtureID)

        if len(match) != 0:
            appointments.extend(match.copy())
        count = count+1
        total = count * perCent

        print("\r",end='')
        print("%.2f" % (total*100), end='')
        print("%", end='')

    # write csv's
    # save the appointments to a csv
    print("Found " + str(len(appointments)-1) + " appointments")

    if appointmentsFile != '':
        with open(appointmentsFile,"w",newline='') as file:
            writer = csv.writer(file)
            writer. writerows(appointments)
        print("Output appointments to csv")

    # output fixtures to CSV
    if fixturesFile != '':
        output = []
        output.append(fixture)
        for f in fixtures:
            output.append(list(f))
        with open(fixturesFile,"w",newline='') as file:
            writer = csv.writer(file)
            writer.writerows(output)
        print ("Output fixtures to CSV")

    # return fixtures and appointments
    if len(appointments) > 1:
        return {'fixturesList':fixtures,'appointmentList':appointments[1:]}
    else:
        return {'fixturesList':fixtures,'appointmentList':[]}

# processAppointments applies the given appointments to schedula
#
# appointments  - list of dicts - {'fixtureID':fid,     # fixture id
#                                  'appointList':[],    # list of appointments in the form {'name':name, 'role':role} where role is one of 'R', 'AR1', 'AR2', 'M', 'A', '4'
#                                   }
# people        - list of names and person id's
#
# returns a list of appointments for the fixtures after updating
def pushAppointments(session, appointments, people, useProxy=False, proxyDict={}):
    
    # appointFixture(session, fixtureID, appointData, proxy=False, proxyDict={})

    updatedAppointments = [] # ['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']

    for match in appointments:
        fixtureID = match['fixtureID']
        appointList = match['appointList']

        appointData = []
        for a in appointList:
            personID = getPersonID(a['name'],people)
            appointData.append([personID, a['role']])

        appointFixture(session, fixtureID, appointData)

        for a in lookupFixture(session, fixtureID):
            updatedAppointments.append(a)
    
    return updatedAppointments


# returns the person id of the name in the list of people, '' is returned for no match
#
# name      - string
# people    - list of {'name':name, 'personID':pid}
def getPersonID(name, people=[]):
    pid = ''

    if name == '':
        return ''

    # attempt exact match
    for p in people:
        if p['name'] == name:
            pid = p['personID']
            break
    
    # attempt close match
    if (pid == ''):
        names = name.split(' ')
        for p in people:
            match = True
            for n in names:
                n = n.replace(',','')
                str1 = n.casefold()
                str2 = p['name'].casefold()
                if (str1 not in str2):
                    match = False
            if match:
                pid = p['personID']

    if pid == '':
        raise Exception('Person: ' + name + 'Not Found')

    return pid
                
# reads a csv of appointments in the form fixtureID,R,AR1,AR2,M,A,4
# returns a list of dicts - {'fixtureID':fid,     # fixture id
#                            'appointList':[],    # list of appointments in the form {'personID':personID, 'role':role} where role is one of 'R', 'AR1', 'AR2', 'M', 'A', '4'}
def readAppointmentList(file='push.csv'):
    dataToPush = readCSV(file)

    appointments = []
    for d in dataToPush:
        fixtureID = d['FixtureID']
        roleList = ['R', 'AR1', 'AR2', 'M', 'A', '4']
        appointList = []
        for r in roleList:
            name = d[r]
            appointList.append({'name':name,'role':r})

        appointments.append({'fixtureID':fixtureID, 'appointList':list(appointList)})

    return appointments



# returns the data in the csv as a list of rows where each row is {heading1:col1, geading2:col2, etc}
def readCSV(filename):
    output = []
    with open(filename, mode='r') as file:
        reader = csv.DictReader(file, delimiter=',')
        for row in reader:
            output.append(row)

    return output


# getOfficials gets all the referees form the given year and pannel and outputs to a csv
#
# session            - session to use for connection to schedula
# year               - calander year to pull from, e.g. '2020'
# pannel             - Only referees on the selected pannel are returned, panel '' will return from all panels
# personsFile        - file name to store people in
# useProxy           - use an http proxy if true
# proxyDict          - address of proxy
#
# A list of dicts is returned in the form {'name':name, 'personID':pid}
def getOfficials(session, year='2020', pannel='', personsFile='People.csv', useProxy=False, proxyDict={}):
    # get organisations
    organisations = getOrganisations(session, useProxy, proxyDict)

    # get seasons
    seasons = []
    for org in organisations:
        orgNum = org[0]
        orgName = org[1]
        s = getSeasons(session, orgNum, useProxy, proxyDict)

        for se in s:
            seasons.append([orgNum, orgName, se[0], se[1]])

    # print('Found Seasons:')
    # #print("['OrgID','Org','SID','SName']")
    # for ss in seasons:
    #     print(ss)

    # select based on year
    if year != '':
        removes = []
        for ss in seasons:
            if ss[3] != year:
                removes.append(ss)

        for r in removes:
            seasons.remove(r)

    
    # get one week from each season
    weeks = []
    for season in seasons:
        seasonID = season[2]
        ws = getSeasonWeeks(session, seasonID, useProxy, proxyDict)

        # get the last week
        if len(ws) != 0:
            w = ws[-1]
            weeks.append([season[0], season[1], season[2], season[3], w[0], w[1]])

    # print('Looking at Weeks:')
    # print("['OrgID','Org','SID','SName','WID','WName']")
    # for week in weeks:
    #     print(week)

    # get one fixture from each week
    fixtures = []
    for week in weeks:

        orgNum = week[0]
        orgName = week[1]
        seasonID = week[2]
        seasonName = week[3]
        weekID = week[4]
        weekName = week[5]

        # get the fixtures for the week
        fixs = getFixturesForWeek(session, seasonID, weekID, useProxy, proxyDict)

        # select the first one
        if len(fixs) > 0:
            f = fixs[0]
            fixtures.append([orgNum,orgName,seasonID,seasonName,weekID,weekName,f[0],f[1],f[2],f[3],f[4],f[5],f[6]])

    
    # get from the selected pannel
    info = []
    print(' Looking at Fixtures:')
    if pannel != '':
        for f in fixtures:
            print(f)
            fixtureID = f[12]
            info.append(getRefInfo(session, fixtureID, pannelName=pannel, proxy=useProxy, proxyDict=proxyDict))

    # get from all pannels
    else:
        for f in fixtures:
            print(f)
            fixtureID = f[12]

            # get list of panels
            initialInfo = getRefInfo(session, fixtureID, proxy=useProxy, proxyDict=proxyDict)
            pannels = initialInfo['pannels']

            # get info from each pannel
            for p in pannels:
                panelName = p[1]
                info.append(getRefInfo(session, fixtureID, pannelName=panelName, proxy=False, proxyDict={}))

            
    # get all the unique person ids
    personIDs = []
    people = []
    for i in info:
        referees = i['referees'] #[name, appointID, personID]
        for r in referees:
            name = r[0]
            appointID = r[1]
            pid = r[2]

            if pid not in personIDs:
                people.append({'name':name, 'personID':pid})

    # output to csv
    out = [['name', 'personID']]
    for p in people:
        out.append([p['name'], p['personID']])
    if personsFile != '':
        with open(personsFile,"w",newline='') as file:
            writer = csv.writer(file)
            writer. writerows(out)

    return people

# lookupFixture returns the fixture details and appointments from schedula
#
# fixtureID - id number of the fixture
# session   - session to use for connection to schedula
#
# Returns a list of appointment details (including confirmation status)
# appointment format: ['fixtureid','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']
def lookupFixture(session, fixtureID):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    r = session.get(url)

    text = r.text
    text = text.split('<')

    tableStart = False

    appointments = []
    
    # ['officialName','officialID','selectedRole','selectedRoleID','acceptStatus']
    appointment = ["","","","",""]

    for t in text:
        #print("Looking at:")
        #print(t)

        # check for end of table
        if t.startswith('/table'):
            #print('End Table')
            break


        if tableStart:
            # check for new row
            if t.startswith('tr'):
                appointment = ["","","","",""]

            # check for data complete
            if appointment[0] != "" and appointment[1] != "" and appointment[2] != ""and appointment[3] != ""and appointment[4] != "":
                #print('Record complete')
                output = appointment.copy()
                output.insert(0,fixtureID)
                appointments.append(output.copy())
                appointment = ["","","","",""]

            # check for name
            if t.startswith('td>'):
                try:
                    data = t.split('>')[1]

                    # record name
                    if appointment[0] == '':
                        appointment[0] = data

                except:
                    #print('Ignore table data')
                    pass
            
            # record appointment type
            elif t.startswith('option value='):
                try:
                    if t.split(' ')[2].startswith('selected'):
                        appointment[2] = t.split('>')[1]
                        appointment[3] = t.split('"')[1]
                except:
                    #print('Ignore')
                    pass

            # record official ID
            elif t.startswith('select'):
                appointment[1] = t.split('xajax_AppointUmpire(')[1].split(',')[0]

            # record acceptance status
            elif t.startswith('img src'):
                try:
                    appointment[4] = t.split('/')[4].split('_')[0]
                except:
                    appointment[4] = '?'

        # check for start of table
        if t.startswith('table'):
            tableStart = True

    return appointments
        

# return list of refs and appointment ids and avalibility status for given fixture
# Also list of appointment types and ids
# returns {'pannels':[id, string], 'appointTypes':[id, string], 'referees':[name, appointID, personID], 'appointments':{'fixtureID:':fixtureID, 'Name':Name, 'appointID':appointID, 'role':role, 'roleID':roleID, 'status':status} }
def getRefInfo(session, fixtureID, pannelName='Referee (built-in)', proxy=False, proxyDict={}):
    # get fixture page
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    r = getPage(session,url,proxy,proxyDict)
    text1 = r.text
    
    # get pannels list
    pant = text1.split('<form name="panels_form">')[1]
    pant = pant.split('<option')

    pannels = []
    for t in pant:
        try:
            value = t.split('alue="')[1].split('"')[0]
            string = t.split('>')[1].split('<')[0]
            pannels.append([value, string])
        except:
            pass

        # check for end of table
        if '</select>' in t:
            break

    # get appointment types list
    appT = text1.split('<form name="appointment_type_form">')[1]
    appT = appT.split('<option')

    appointTypes = []
    for t in appT:
        try:
            value = t.split('alue="')[1].split('"')[0]
            string = t.split('>')[1].split('<')[0]
            appointTypes.append([value, string])
        except:
            pass

        # check for end of table
        if '</select>' in t:
            break

    # get all referees on pannel
    # get referees from the pannel
    # select pannel
    pannel = []
    for p in pannels:
        if p[1] == pannelName:
            pannel = p
            break

    r = changePanel(session, pannel[0], fixtureID, proxy=proxy, proxyDict=proxyDict)
    text = r.text

    referees = getRefsFromText(text)

    # get exsisting appointments
    #appointments = lookupFixture(session, fixtureID)
    text = text1
    text = text.split('<')

    tableStart = False

    appointments = []
    
    # ['officialName','officialID','selectedRole','selectedRoleID','acceptStatus']
    appointment = ["","","","",""]

    for t in text:
        #print("Looking at:")
        #print(t)

        # check for end of table
        if t.startswith('/table'):
            #print('End Table')
            break


        if tableStart:
            # check for new row
            if t.startswith('tr'):
                appointment = ["","","","",""]

            # check for data complete
            if appointment[0] != "" and appointment[1] != "" and appointment[2] != ""and appointment[3] != ""and appointment[4] != "":
                #print('Record complete')
                output = appointment.copy()
                output.insert(0,fixtureID)
                appointments.append(output.copy())
                appointment = ["","","","",""]

            # check for name
            if t.startswith('td>'):
                try:
                    data = t.split('>')[1]

                    # record name
                    if appointment[0] == '':
                        appointment[0] = data

                except:
                    #print('Ignore table data')
                    pass
            
            # record appointment type
            elif t.startswith('option value='):
                try:
                    if t.split(' ')[2].startswith('selected'):
                        appointment[2] = t.split('>')[1]
                        appointment[3] = t.split('"')[1]
                except:
                    #print('Ignore')
                    pass

            # record official ID
            elif t.startswith('select'):
                appointment[1] = t.split('xajax_AppointUmpire(')[1].split(',')[0]

            # record acceptance status
            elif t.startswith('img src'):
                try:
                    appointment[4] = t.split('/')[4].split('_')[0]
                except:
                    appointment[4] = '?'

        # check for start of table
        if t.startswith('table'):
            tableStart = True

    # convert to dict
    e = []
    for a in appointments:
        e.append({'fixtureID:':a[0], 'Name':a[1], 'appointID':a[2], 'role':a[3], 'roleID':a[4], 'status':a[5]})

    # add appointed referees to list
    if len(appointments) != 0:
        for a in e:
            for ref in referees:
                if ref[0] == a['Name']:
                    ref[1] = a['appointID']
                    #print('Ammended:')
                    #print(ref)
                    break

    return {'pannels':pannels, 'appointTypes':appointTypes, 'referees':referees, 'appointments':e }


# Appoints the given refids and appointment types to the given match
#
# session            - session to use for connection to schedula
# fixtureID          - id number of the fixture
# appointData        - [refId, appointmentTypeIds]
# refId              - person id of the referees to be appointed
# appointmentTypeId  - id of the appointment type, i.e. R,AR1,AR2,M,A,4,N (referee, ar1, ar2, mentor, assessor, 4th official, Null(remove them))
#
# Aborts on failure.
def appointFixture(session, fixtureID, appointData, proxy=False, proxyDict={}):
    try:
        print('\nModify Fixture <' + fixtureID + '>...')

        # get info about the fixutre
        appointInfo = getRefInfo(session, fixtureID, proxy=proxy, proxyDict=proxyDict)
        pannels = appointInfo['pannels']
        appointTypes = appointInfo['appointTypes']
        referees = appointInfo['referees']
        exsistingAppointments = appointInfo['appointments']

        # get the selected referees
        appointIds = [] # [name, appoint id, person id, appointType, appointmentTypeid]
        for pid in appointData:
            if pid[0] == '':
                appointIds.append(['','','',pid[1], ''])
            else:
                foundRef = False
                for ref in referees:
                    # check for unknown appointment id, i.e. Referee already appointed
                    if ref[1] == '':
                        print(" Fialed to appoint " + ref[0] + " to " + fixtureID)
                        foundRef = True
                        break
                    if ref[2] == pid[0]:
                        appointIds.append([ref[0], ref[1], ref[2], pid[1], ''])
                        foundRef = True
                        break
                if not foundRef:
                    raise Exception('Unknown Referee:' + str(pid))

        # get the appointment ids for each type r,ar1,ar2,m,a,4
        R = []
        AR1 = []
        AR2 = []
        M = []
        A = []
        R4 = []

        pAR1 = re.compile('A.*R.*1') # A*R*1* = AR1
        pAR2 = re.compile('A.*R.*2') # A*R*2* = AR2
        pM = re.compile('.*Mentor.*') # *Mentor* = M
        pA = re.compile('R.*Assessor') # R*Assessor = A
        pR = re.compile('Referee') # R* = R
        p4 = re.compile('.*4.*') # *4* = 4

        for a in appointTypes:
            if (pAR1.match(a[1]) != None):
                AR1 = a
            elif pAR2.match(a[1]) != None:
                AR2 = a
            elif pM.match(a[1]) != None:
                M = a
            elif pA.match(a[1]) != None:
                A = a
            elif a[1] == 'Referee':
                R = a
            elif p4.match(a[1]) != None:
                R4 = a
                
        typeIds = {'R':R, 'AR1':AR1, 'AR2':AR2, 'M':M, 'A':A, '4':R4, 'N':['-1','Unappoint']}

        # add appointment type
        for a in appointIds:
            aType = a[3]
            a[4] = typeIds[aType]


        # check against exsisting appointments
        print('  Removing appointments:')
        removedExsisting = []
        removedNew = []
        if len(exsistingAppointments) != 0:
            for e in exsistingAppointments:
                # remove appointment if a role or person is already appointed differently
                for a in appointIds:
                    if ((e['roleID'] == a[4][0]) and (e['Name'] != a[0])) or ((e['Name'] == a[0]) and (e['roleID'] != a[4][0])):
                        # check not already removed
                        if not (e in removedExsisting):
                            print('   ' + str(e))
                            r = UnappointUmpire(session, e['appointID'], e['roleID'], fixtureID, proxy=proxy, proxyDict=proxyDict)
                            removedExsisting.append(e)
                            text = r.text
                            # update appoint ids
                            referees = getRefsFromText(text)
                            for a in appointIds:
                                for r in referees:
                                    if a[2] == r[2]:
                                        a[1] = r[1]
                                        break
                            # TODO: update ids in exsistingAppointments? OR work out sonthing better for getting appointment ids
                
                    # do not appoint if the appointment already exsists
                    if ((e['roleID'] == a[4][0]) and (e['Name'] == a[0])):
                        removedNew.append(a)
                        
        print('  Adding appointments:')
        #print(['name', 'appoint id', 'person id', 'appointType', '[appointmentTypeid, name]'])

        # Make appointments
        for a in appointIds:
            if a in removedNew:
                print('   ' + str(a) + 'Already in schedula, skipping...')
            elif (a[3] != 'N') and (a[1] != ''):
                print('   ' + str(a))
                #r = changeAppointmentType(session, typeIds['4'][0], fixtureID, proxy=proxy, proxyDict=proxyDict)
                r = changeAppointmentType(session, a[4][0], fixtureID, proxy=proxy, proxyDict=proxyDict)
                r = AppointUmpire(session,a[1],a[4][0],fixtureID, proxy=proxy, proxyDict=proxyDict)

        # save changes
        SaveAppointments(session, fixtureID)

        print('Modify <' + fixtureID + '> Complete.')
    
    # an error has occured, bachtrack changes
    except Exception as ex:
        try:
            exc_info = sys.exc_info()

            DiscardChanges(session, fixtureID, proxy=proxy, proxyDict=proxyDict)
            JustClose(session, fixtureID, proxy=proxy, proxyDict=proxyDict)

            print('Modify <' + fixtureID + '> Cancled. Error information below:')

            traceback.print_exception(*exc_info)
        finally:
            del exc_info


# yep
def getRefsFromText(text):
    refT = text.split('<td>')
    
    # debug
    # if 'Pending appointments will also appear above.' in text:
    #     print('has extra')
    #     fileOut = open('temp.txt','w')
    #     for t in refT:
    #         fileOut.write(t + '\n')
    #     fileOut.close()

    referees = []
    table2 = False
    for t in refT:
        if table2:
            # check for removed
            if ('value="Removed"' in t) or ('<img src' in t):
                pass
            else:
                n = t.split('</td')
                if len(n) == 2:
                    name = n[0]
                else:
                    appointID = t.split('AppointUmpire(')[1].split(',')[0]

                    # update
                    for r in referees:
                        if r[0] == name:
                            #print('ammended ' + str(r))
                            r[1] = appointID
                            #print('      to ' + str(r))
                            appointID = ''
                            personID = ''
                            name = ''
                            break

        else:
            try:
                name = t.split('<b>')[1].split('</b>')[0]
                appointID = t.split('AppointUmpire(')[1].split(',')[0]
            except:
                try:
                    personID = t.split('personid=')[1].split('&')[0]
                    referees.append([name, appointID, personID])
                    # clear fileds, this prevents duplicates
                    appointID = ''
                    personID = ''
                    name = ''
                except:
                    pass

            # check for end of table, start of new one
            if '<![CDATA[S<table>' in t:
                table2 = True

    # tracking
    # print('\n\n')
    # for ref in referees:
    #     if (ref[2] == '11504519') or (ref[2] == '13957370'):
    #         print(ref)
    
    return referees


# returns unix time in ms
def getXjxr():
    t = int(round(time.time()*1000))
    return str(t)

# converts date in 2020-02-03 format into a date object
def stringToDate(dateString):
    subStrings = dateString.split('-')
    year = int(subStrings[0])
    month = int(subStrings[1])
    day = int(subStrings[2])

    return date(year,month,day)

# returns true if the given competition is blacklisted
def isBlacklisted(competitionString):
    ages = ['6','7','8','9','10','11']

    # U6, U7, ...
    for a in ages:
        pattern = re.compile('U' + a + '.*')
        if pattern.match(competitionString) != None:
            #print('Blacklist Hit <' + competitionString + '>')
            return True

    # Under 6, Under 7, ...
    for a in ages:
        pattern = re.compile('Under.' + a + '.*')
        if pattern.match(competitionString) != None:
            #print('Blacklist Hit <' + competitionString + '>')
            return True
    
    return False

# converts a role string into a consistant designator
# e.g. Assistant Referee 1-> AR1
#      AR 1 -> AR1
def roleStringToLetter(role):
    # patterns for matching role names
    pAR1 = re.compile('A.*R.*1') # A*R*1* = AR1
    pAR2 = re.compile('A.*R.*2') # A*R*2* = AR2
    pM = re.compile('.*Mentor.*') # *Mentor* = M
    pA = re.compile('R.*Assessor') # R*Assessor = A
    pR = re.compile('Referee') # R* = R
    p4 = re.compile('.*4.*') # *4* = 4

    # output variable
    roleName = ''

    # match role name to one of R,AR1,AR2,M,A,R4
    if (pAR1.match(role) != None):
        roleName = 'AR1'
    elif pAR2.match(role) != None:
        roleName = 'AR2'
    elif pM.match(role) != None:
        roleName = 'M'
    elif pA.match(role) != None:
        roleName = 'A'
    elif role == 'Referee':
        roleName = 'R'
    elif p4.match(role) != None:
        roleName = '4'
    else:
        roleName = 'Other'

    return roleName

################################
## Schedula xjx api functions ##
################################
def changePanel(session, panelId, fixtureID, onlyAvaliable='false', proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=ChangePanel&xjxr=' + getXjxr() + '&xjxargs[]=N' + panelId + '&xjxargs[]=N' + fixtureID + '&xjxargs[]=B'+ onlyAvaliable
    print('    changePanel: ' + data)
    return post(session, url, data, proxy, proxyDict)

def changeAppointmentType(session, appointmentId, fixtureID, proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=ChangeAppointmentType&xjxr=' + getXjxr() + '&xjxargs[]=S' + appointmentId + '&xjxargs[]=N' + fixtureID
    print('    changeAppointmentType: ' + data)
    return post(session, url, data, proxy, proxyDict)

def AppointUmpire(session, refAppointId, appointmentType, fixtureID, arg = 'false', proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=AppointUmpire&xjxr=' + getXjxr() + '&xjxargs[]=N' + refAppointId + '&xjxargs[]=S' + appointmentType + '&xjxargs[]=N' + fixtureID + '&xjxargs[]=B' + arg
    print('    AppointUmpire: ' + data)
    return post(session, url, data, proxy, proxyDict)

def DisplayOnWeb(session, refAppointid, appointmentType, fixtureID, flag, proxy=False, proxyDict={}): #flag=0 to display on website
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=DisplayOnWeb&xjxr=' + getXjxr() + '&xjxargs[]=N' + refAppointid + '&xjxargs[]=N' + appointmentType + '&xjxargs[]=N' + fixtureID + '&xjxargs[]=N' + flag
    print('    DisplayOnWeb: ' + data)
    return post(session, url, data, proxy, proxyDict)

def UnappointUmpire(session, refappointid, appointType, fixtureID, proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=UnappointUmpire&xjxr=' + getXjxr() + '&xjxargs[]=N' + refappointid + '&xjxargs[]=N' + appointType + '&xjxargs[]=N' + fixtureID
    print('    UnappointUmpire: ' + data)
    return post(session, url, data, proxy, proxyDict)

def SaveAppointments(session, fixtureID, proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=SaveAppointments&xjxr=' + getXjxr() + '&xjxargs[]=N' + fixtureID
    print('    SaveAppointments: ' + data)
    return post(session, url, data, proxy, proxyDict)

def DiscardChanges(session, fixtureID, arg1='false', arg2='false', proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=DiscardChanges&xjxr=' + getXjxr() + '&xjxargs[]=N' + fixtureID + '&xjxargs[]=B' + arg1 + '&xjxargs[]=B' + arg2
    print('    DiscardChanges: ' + data)
    return post(session, url, data, proxy, proxyDict)

def JustClose(session, fixtureID, proxy=False, proxyDict={}):
    url = "https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=" + fixtureID + "&skeleton=true"
    data = 'xjxfun=JustClose&xjxr=' + getXjxr() + '&xjxargs[]=N' + fixtureID
    print('    JustClose: ' + data)
    r = post(session, url, data, proxy, proxyDict)
    text = r.text
    t = text.split('func="')[1].split('"')[0]

    if t == ("confirmClose(" + fixtureID + ")"):
        print('Bad Close')
        raise Exception("Bad Close")
        DiscardChanges(session, fixtureID, proxy=proxy, proxyDict=proxyDict)

    return r
