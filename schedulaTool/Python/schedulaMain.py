# command line program to interface with schedula

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

import sys
import getopt
import getpass
import time
import re
import csv
import traceback
from datetime import date
import schedulaInterface as schedula


# print command line program usage
def usage():
    print("Schedula Command Line Tool\n Use with care, actions cannot be undone.\n It is the responsibility of the user to take due care using this tool and to ensure they have authorisation for all actions.\n While care has been taken to ensure this tool works correctly, no gaurentees are made and the author accept no responsibility for unintended consequences of using this tool.\n The author of this tool has no control over schedula which may change how it works with no notice.\n This tool is provided under the MIT License, see LICENCE.txt for details.\n")
    print("Usage: python " + __file__ + " command [options] [arguments]")
    print("\nCommands:")
    print(" pullAll     Gets all fixtures and appointments from schedula. Use -s to specify a season. Use -o to do pullP at the same time")
    print(" pullN       Gets all fixtures and appointments for the next 28 days. use -n to change the number of days, use -N to set the start date in the form yyyy-mm-dd")
    print(" pullP       Gets all the match officials from shcedula, exports to the file given by -o")
    print(" push        Pushes the appointments in the file given by -i to schedula, using the file given by -o. Checks appointments have not changed compared to the file specified by -f.")
    print("\nOptions:\n -f   filename\n -s   season (e.g. 2020)\n -u   username\n -p   password\n -i   File of fixtures. Used with command \"push\".\n -o   File of officials. Used with commands \"pullP\", \"push\" or \"pullAll\".\n -x   HTTP proxy address (e.g. localhost:8080)\n -n   Number of days to pull. Used with command \"pullN\"\n -N   Start date, used with command \"pullN\". Must be in the form yyyy-mm-dd\n -h   Display usage")

# pullP command
# Gets all the names and person Ids from schedula
def pullP(session, season, peopleFile, useProxy, proxyDict):
    print('Command: pullP')
    print(' peopleFile: ' + peopleFile)
    print(' season: ' + season)
    p = schedula.getOfficials(session, year=season, pannel='', personsFile=peopleFile, useProxy=useProxy, proxyDict=proxyDict)

# pullAll command
# Gets all fixtures and appointments for the given season. If season is '' all avaliable records are pulled
def pullAll(session, filename, season, useProxy, proxyDict):
    print('Command: pullAll')
    print(' filename: ' + filename)
    print(' season: ' + season)

    # pull data from schedula
    data = schedula.pullAll(session, year=season, fixturesFile='', appointmentsFile='', useProxy=False, proxyDict={})
    
    if data is not None:
        fixtures = data['fixturesList']
        appointments = data['appointmentList']

        # output to csv
        if filename != '':
            writeToCsv(fixtures, appointments, filename)

# pullN command
# Gets all fixtures and appointments form the start date plus N days. If no start date specified, the current date is used
def pullN(session, filename, season, startDay, N, useProxy, proxyDict):
    print('Command: pullN')
    print(' filename: ' + filename)
    print(' season: ' + season)
    print(' startDay: ' + str(startDay))
    print(' N:' + str(N))

    # pull data from schedula
    data = schedula.update28(session, season, fixturesFile='', appointmentsFile='', startDate=startDay, numberDays=N, useProxy=False, proxyDict={})
    fixtures = data['fixturesList']
    appointments = data['appointmentList']

    # output to csv
    if filename != '':
        writeToCsv(fixtures, appointments, filename)

# push command
# Gets the appointments in the given push file and pushes them to schedula
# Checkes for changes in schedula and does not appoint clashes
def push(session, filename, pushFile, peopleFile, useProxy, proxyDict):
    print('Command: push')

    # Load filename - ['FixtureID','OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date', 'day','Time','Home','Away','Ground','Referee','AR1','AR2','Mentor','Assessor','4th Official','Other','Status']
    print(' filename: ' + filename)
    storedAppointments = schedula.readCSV(filename)

    # Load pushFile - ['FixtureID','R','AR1','AR2','M','A','4']
    print(' pushFile: ' + pushFile)
    pushData = schedula.readAppointmentList(pushFile)

    # Load people file
    print(' peopleFile: ' + peopleFile)
    people = schedula.readCSV(peopleFile)

    # For each fixtureID in pushFile: check againsed schedula for changes, if change found does it clash with appointment? if yes don't apoint.
    clashes= []
    for d in pushData:
        # Extract fixture ID
        fixtureID = d['fixtureID']

        # get current fixture status
        serverAppoints = schedula.lookupFixture(session, fixtureID) # ['fixtureid','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']
        roles = ['R', 'AR1', 'AR2', 'M', 'A', '4', 'Other']
        serverRoles = {'R':'', 'AR1':'', 'AR2':'', 'M':'', 'A':'', '4':'', 'Other':''}
        
        # fit names to roles
        for a in serverAppoints:
            role = schedula.roleStringToLetter(a[3])
            serverRoles[role] = a[1]

        # get stored data
        storedFixture = {}
        for s in storedAppointments:
            if s['FixtureID'] == fixtureID:
                storedFixture = s
                break

        # fixture not stored, reject
        if not storedFixture:
            clashes.append(d)

        # check status
        elif not (storedFixture['Status'] == 'ok'):
            clashes.append(d)

        else:
            # get stored roles
            storedRoles = {'R':storedFixture['Referee'], 'AR1':storedFixture['AR1'], 'AR2':storedFixture['AR2'], 'M':storedFixture['Mentor'], 'A':storedFixture['Assessor'], '4':storedFixture['4th Official'], 'Other':storedFixture['Other']}

            # Check sever and stored data matches, if not reject change
            for r in roles:
                if serverRoles[r] != storedRoles[r]:
                    clashes.append(d)
                    break

    # remove clashes
    for c in clashes:
        print(' Cannot appoint fixture:' + str(c))
        pushData.remove(c)

    # display changes that are about to occur
    print("Changes are about to be applied, this action cannot be undone")
    # for p in pushData:
    #    print(p) # prints all changes

    # HOLD POINT - require user acknowledgment to continue
    hold = True
    strIn = input('Please Confirm (yes or no):')
    while hold:
        if strIn == 'yes':
            hold = False

        if strIn == 'no':
            print('Abort, user permission denied')
            sys.exit(0)

        if strIn not in ['yes', 'no']:
            strIn = input("Please enter 'yes' or 'no':")
            if (strIn not in ['yes', 'no']) or (strIn == 'no'):
                print('Abort, user permission not granted')
                sys.exit(2)

    # For each ficture in pushfile: appoint fixture
    updated = schedula.pushAppointments(session,pushData,people,useProxy=useProxy, proxyDict=proxyDict)

    # TODO: update filename


# function to write fixtures and appointments to a csv
#       fixtures     - ['OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date','Time','Home','Away','Ground','FixtureID']
#       appointments - ['fixtureID','officialName','appointID','selectedRole','selectedRoleID','acceptStatus']
def writeToCsv(fixtures, appointments, filename):
    print(' Write to csv <' + filename + '>')

    # output titles
    headerRow = ['FixtureID','OrgID','OrgName','SeasonID','SeasonName','WeekID','WeekName','Competition','Date', 'day','Time','Home','Away','Ground','Referee','AR1','AR2','Mentor','Assessor','4th Official','Other','Status','Rstatus','AR1status','AR2status','Mstatus','Astatus','4status']
    output = []
    # output.append(headerRow)

    # Apply format
    for fix in fixtures:
        fixID = fix[12]
        roles = {'R':'', 'AR1':'', 'AR2':'', 'M':'', 'A':'', '4':'', 'Other':''}
        status = {'R':'', 'AR1':'', 'AR2':'', 'M':'', 'A':'', '4':'', 'Other':''}
        statusFlag = 'ok'

        fixAppoints = []
        # get appointments for this fixture
        for a in appointments:
            if a[0] == fixID:
                fixAppoints.append(a)

        # assign roles
        for a in fixAppoints:
            roleName = a[3]

            # match role name to one of R,AR1,AR2,M,A,R4
            roleName = schedula.roleStringToLetter(roleName)

            # fill the role
            if roles[roleName] == '':
                roles[roleName] = a[1] # referee name
                status[roleName] = a[5] # accept status
            else:
                statusFlag = 'Appointment Error, multiple appointments to ' + roleName
                print('Warning: multiple appointments to ' + roleName + '. FixtureID: ' + fixID + '. See: https://schedula.sportstg.com/index.php?action=admin/appointments/appoint_match&fixtureid=' + fixID)

        # convert date to somthing nicer (before: Sat Jun 27) (after: 27-Jun-2020)
        dateStrs = fix[7].split(' ')
        yearStr = fix[3]
        niceDate = dateStrs[2] + '-' + dateStrs[1] + '-' + yearStr

        # create the row of data
        output.append([fixID,fix[0],fix[1],fix[2],fix[3],fix[4],fix[5],fix[6],niceDate,dateStrs[0],fix[8],fix[9],fix[10],fix[11],roles['R'],roles['AR1'],roles['AR2'],roles['M'],roles['A'],roles['4'],roles['Other'],statusFlag,status['R'],status['AR1'],status['AR2'],status['M'],status['A'],status['4']])
        
    getFixID = lambda a : a[0]

    # sort by fixtureID
    output.sort(key=getFixID)

    # output to csv
    with open(filename,"w",newline='') as file:
        writer = csv.writer(file)
        writer.writerows([headerRow])
        writer.writerows(output)

# entry point, parse commandline arguments and call appropriate method
def main(argv):
    # parameters
    command = ''
    filename = ''
    season = ''
    username = ''
    password = ''
    pushFile = ''
    peopleFlag = False
    peopleFile = 'people.csv'
    numDaysToPull = 28
    startDate = date.today()
    useProxy = False
    proxyDict = {}

    # get commandline options
    try:
        opts, args = getopt.gnu_getopt(argv,"f:s:u:p:i:x:n:o:N:rh")
    except getopt.GetoptError as err:
        print(str(err))
        usage()
        sys.exit(2)

    # process the argument(s)
    if len(args) != 1:
        print('Error: require exactly one command, use \"help\" for more info')
        sys.exit(2)
    else:
        command = args[0]
        if command not in ['pullAll', 'pullN', 'push', 'pullP', 'help']:
            print('Unknown command: ' + command)
            sys.exit(2)
        
        if command == 'help':
            usage()
            sys.exit(0)

    # process the options
    for opt, arg in opts:
        if opt == '-h':
            usage()
            sys.exit()
        elif opt == '-f':
            filename = arg
        elif opt == '-s':
            season = arg
        elif opt == '-u':
            username = arg
        elif opt == '-p':
            password = arg
        elif opt == '-i':
            if command != 'push':
                print("-i can only be used with command \"push\"")
                sys.exit(2)
            pushFile = arg
        elif opt == '-x':
            useProxy = True
            httpProxy = "http:" + arg
            httpsProxy = "https://" + arg
            proxyDict = {"http":httpProxy, "https":httpsProxy}
        elif opt == '-n':
            if command != 'pullN':
                print("-n can only be used with command \"pullN\"")
                sys.exit(2)
            try:
                numDaysToPull = int(arg)
            except:
                print('-n must specify an integer number of days (e.g. -n28)')
                sys.exit(2)
        elif opt == '-N':
            if command != 'pullN':
                print("-N can only be used with command \"pullN\"")
                sys.exit(2)
            try:
                strs = arg.split('-')
                startDate = date(int(strs[0]),int(strs[1]),int(strs[2]))
            except:
                print('-N, format must be yyyy-mm-dd')
                sys.exit(2)
        elif opt == '-o':
            if command not in ['pullP', 'push', 'pullAll']:
                print("-o can only be used with commands \"pullN\", \"push\" or \"pullAll\"")
                sys.exit(2)
            peopleFile = arg
            peopleFlag = True

    # get schedula login details
    if username == '':
        print('Enter schedula login details')
        username = input('Email:')
    if password == '':
        password = getpass.getpass('Password:')

    # login to schedula
    session = schedula.getSession(username, password, useProxy, proxyDict)

    # process the command
    if command == 'pullAll':
        pullAll(session, filename, season, useProxy, proxyDict)
        if peopleFlag:
            pullP(session, season, peopleFile, useProxy, proxyDict)
    elif command == 'pullN':
        pullN(session, filename, season, startDate, numDaysToPull, useProxy, proxyDict)
    elif command == 'push':
        push(session, filename, pushFile, peopleFile, useProxy, proxyDict)
        i = input("Press enter to exit")
    elif command == 'pullP':
        pullP(session, season, peopleFile, useProxy, proxyDict)


if __name__ == '__main__':
    try:
        main(sys.argv[1:])
    except:
        try:
            exc_info = sys.exc_info()
            traceback.print_exception(*exc_info)

            print("Error occured. Arguments:")
            print(sys.argv)
            i = input("Press enter to exit")

        finally:
            del exc_info