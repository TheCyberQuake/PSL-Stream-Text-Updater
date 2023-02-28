# coding: utf-8
# Use utf-8 encoding to allow non-ascii characters
# We only NEED this because people want to use special characters to drive me insane



# Import a bunch of stuff. Just ALL the stuff. All of it.
from __future__ import print_function
import sys
import pickle
import os.path
import webbrowser
import shutil
import time
import os
import re
import codecs


# import only system from os
from os import system, name
  
# import sleep to show output for some time period 
from time import sleep 

# Install required Google API modules
import importlib, os

# Function to auto-magically install missing modules
def import_neccessary_modules(modname:str)->None:
    '''
        Import a Module,
        and if that fails, try to use the Command Window PIP.exe to install it,
        if that fails, because PIP in not in the Path,
        try find the location of PIP.exe and again attempt to install from the Command Window.
    '''
    try:
        # If Module it is already installed, try to Import it
        importlib.import_module(modname)
        print(f"Importing {modname}")
    except ImportError:
        # Error if Module is not installed Yet,  the '\033[93m' is just code to print in certain colors
        print(f"\033[93mModule [{modname}] is missing. Installing...\033[0m")
        #print("I will need to install it using Python's PIP.exe command.\033[0m")
        if os.system('PIP --version') == 0:
            # No error from running PIP in the Command Window, therefor PIP.exe is in the %PATH%
            os.system(f'PIP install {modname}')
        else:
            # Error, PIP.exe is NOT in the Path!! So I'll try to find it.
            pip_location_attempt_1 = sys.executable.replace("python.exe", "") + "pip.exe"
            pip_location_attempt_2 = sys.executable.replace("python.exe", "") + "scripts\pip.exe"
            if os.path.exists(pip_location_attempt_1):
                # The Attempt #1 File exists!!!
                os.system(pip_location_attempt_1 + " install " + modname)
            elif os.path.exists(pip_location_attempt_2):
                # The Attempt #2 File exists!!!
                os.system(pip_location_attempt_2 + " install " + modname)
            else:
                # Neither Attempts found the PIP.exe file, So i Fail...
                #print(f"\033[91mAbort!!!  I can't find PIP.exe program!")
                print(f"You'll need to manually install the Module: {modname} in order for this program to work.")
                print(f"Find the PIP.exe file on your computer and in the CMD Command window...")
                print(f"   in that directory, type    PIP.exe install {modname}\033[0m")
                exit()

# Make sure google API modules are installed, if not, install them
import_neccessary_modules('google-api-python-client')
import_neccessary_modules('google-auth-httplib2')
import_neccessary_modules('google-auth-oauthlib')
import_neccessary_modules('requests')

# GoogleAPI gubbins to access google sheets
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests


# A clear function that will work across OS
def clear(): 
  
    # for windows 
    if name == 'nt': 
        _ = system('cls') 

    # for mac and linux(here, os.name is 'posix') 
    else: 
        _ = system('clear') 


# If modifying these scopes, delete the file token.pickle.
# Grants Google API access to read spreadsheets, used to pull data from PSL spreadsheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


# Adds logging functionality
def log(message = 'Logging Error', notime = False):
    
    # Gets local time
    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
    # Opens file for writing
    log = codecs.open('log.txt','a','utf-8')
    #Write log entry, close file
    if notime:
        # Write log entry without a timestamp
        log.write(str(message) + '\n')
    else:
        # Write log entry with a timestamp
        log.write(current_time + "  " + str(message) + '\n')
    log.close()

# Print and log error
def errinfo():
    print(sys.exc_info()[0])
    log(sys.exc_info()[1])

# Exit script with error
def errexit(errcode = 1, errinfo = ''):
    log('!!!!!Script Unsuccessful(' + str(errcode) + ')!!!!!', True)
    sys.exit(errcode)

# Try to access a google sheet, error handle if it can't
def trysheet(service, SHEETID, SHEETRNG, ERRMSG, ERRCODE, MJRDEM = 'ROWS'):
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SHEETID,
                    range=SHEETRNG,
                    majorDimension=MJRDEM).execute()
        values = result.get('values', [])
        
        if not values:
            print(ERRMSG)
            log(ERRMSG)
            print('')
            print('Press Enter to exit')
            input()
            errexit(ERRCODE)
        return values
    except:
        print(ERRMSG)
        errinfo()
        print('')
        print('Press Enter to exit')
        input()
        errexit(ERRCODE)

# used for text file output
def outfile(fileLoc, fileText):
    try:
        # Open file for write with utf-8 encoding because players want to be special and add 'fancy' characters ¯\_(ツ)_/¯
        WriteFile = codecs.open(fileLoc,'w','utf-8')
        WriteFile.write(fileText)
        return True
    except:
        # Whoopsies, something went wrong. Print to screen and log
        log('Error: Write to ' + fileLoc + ' was unsuccessful!')
        print('Error: Write to ' + fileLoc + ' was unsuccessful!')
        errinfo()
        return False
    finally:
        # Close file no matter if the write was successful or not
        WriteFile.close()

def main():

    
    currentVer = 1.3
    latestVer = None
    _startup_cwd = os.getcwd()
    
    # Check for updates
    try:
        # Query github API to grab latest version tag for latest version number
        response = requests.get('https://api.github.com/repos/TheCyberQuake/PSL-Stream-Text-Updater/releases/latest')
        latestVer = float(response.json()["tag_name"])
        log('Current version: ' + str(currentVer))
        log('Latest version: ' + str(latestVer))
        # See if we are out of date
        if currentVer < latestVer:
            log('Newer version detected, attempting to update...')
            log('Downloading latest release')
            # Attempt to download and overwrite script
            # TODO: make script name dynamic to current file name in case user changes it
            myfile = requests.get('https://github.com/TheCyberQuake/PSL-Stream-Text-Updater/releases/download/' + str(latestVer) + '/updatetext.py')
            open('updatetext.py', 'wb').write(myfile.content)
            log('Downloaded. Attempting to restart into new script...')
            # Attempt to restart script
            os.chdir(_startup_cwd)
            os.execv(sys.executable, ['python'] + sys.argv)
        else:
            log('Running latest version')
    except:
        # Treat this as a non-fatal, log, and show brief message if new version was found but just couldn't download/boot
        if latestVer:
            log('Version ' + str(latestVer) + ' update detected but unable to download')
            print('Latest version ' + str(latestVer) + ' is currently available, but was unable to be downloaded.')
            print('Non-fatal error, will continue in 3 seconds')
            sleep(3)
        else:
            log('Unable to get latest version')
    
    """
    Takes input for name of Team 1 and current week, then
    automatically pulls data from google sheets for team 2 name,
    team players, and team stats, and then outputs to text files
    for use in OBS.
    """
    clear()
    log('', True)
    log('~~~~~Script Start~~~~~', True)

    # Check for API Creds
    if not os.path.exists('credentials.json'):
        print('You will need to obtain a Google API OAuth Client ID named "credentials.json" located in the same folder as this python file')
        print('Reach out to TheCyberQuake to obtain details')
        log('!!!!! Error: Missing creds !!!!!')
        sleep(3)
        errexit(11)

    creds = None

    # Taken straight from updated Kickstart sheets api documenation
    # Will see if I can later get it to use a different oauth screen that doesn't show as "possibly unsafe"
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        errinfo()
        log('!!!!!Error: Could not build Google Sheets API service!!!!!')
        sleep(5)
        errexit(10)
        
    
    """
    This script currently obtains variables from an initial spreadsheet in order to obtain
    the document ID and ranges for the reference sheets to pull PSL data from. The main purposes
    of doing this is the possibility of not needing to update the program in order to get this
    script working for a new season. An override can be created by creating an override.txt in the
    same folder as this script, first line the spreadsheet id, second the ranges to read data from.
    The default sheet this reads from is linked below:
    https://docs.google.com/spreadsheets/d/10Bir5luBjxfR9HNc7e3NRD-1k6xvcfKkciY9xeNbffw/edit?usp=sharing
    """

    # TODO: Look into seeing if this first section to handle grabbing variables can be improved.
    # There's probably a better way to do this
    
    # Check if override file exists, use it if it does
    # TODO: Change override from being another google sheet to instead being a CSV file
    #     Maybe some day I'll get around to this. Maybe by Season 6? Lol that didn't happen
    if os.path.exists('override.txt'):
        print('Override File Exists, using that for variables sheet ID and range')
        log('Override file detected, using it for variables sheet ID and range')
        readfile = open('override.txt','r')
        VARIABLES_SPREADSHEET_ID = readfile.readline().rstrip()
        VARIABLES_RANGE_NAME = readfile.readline().rstrip()
        overridebool = True
        readfile.close()
    else:
        # Use the default variables sheet if no override file exists
        log('No override file exists, using defaults for variable sheet ID and range')
        VARIABLES_SPREADSHEET_ID = '10Bir5luBjxfR9HNc7e3NRD-1k6xvcfKkciY9xeNbffw'
        VARIABLES_RANGE_NAME = 'Sheet1!B1:B11'
        overridebool = False


    # Call Sheet API, pull variables for what sheet IDs and ranges to use
    # This allows spreadsheet IDs and ranges to be updated for new seasons without requiring a client update
    try:
        sheet = service.spreadsheets()
        sleep(1)
        result = sheet.values().get(spreadsheetId=VARIABLES_SPREADSHEET_ID,
                                    range=VARIABLES_RANGE_NAME,
                                    majorDimension='COLUMNS').execute()
        values = result.get('values', [])
    except:
        # Default variable sheet failed, ask user if they want to create an override
        print("Error occured while pulling data from central Variables sheet")
        log('Error occured while pulling data from Variables sheet, asking user for override')
        errinfo()
        print('')
        userconfirm = input('Would you like to create an override that points to another sheet? (Y/n): ')
        if userconfirm.lower() == 'y':
            valid = False
            while not valid:
                clear()
                tempID = input('What is the ID of the sheet you wish to use: ')
                tempRNG = input('What is the Sheet and Range of the variables (Sheet!A1:Z26): ')
                if tempID != '' and tempRNG != '':
                    clear()
                    print('Testing override sheet...')
                    log('Obtained override info. Testing override validity')
                    sheet = service.spreadsheets()
                    result = sheet.values().get(spreadsheetId=tempID,
                                range=tempRNG,
                                majorDimension='COLUMNS').execute()
                    values = result.get('values', [])
                    
                    if not values:
                        print('Error: Invalid Sheet ID or Range')
                        log('Error: Override info invalid')
                    else:
                        log('Override info valid')
                        VARIABLES_SPREADSHEET_ID = tempID
                        VARIABLES_RANGE_NAME = tempRNG
                        valid = True
                        if os.path.exists('override.txt'):
                            os.remove('override.txt')
                        open('override.txt','a').close
                        override = open('override.txt','a')
                        override.write(tempID + '\n')
                        override.write(tempRNG)
                        override.close()
                        print('Variables being populated from override variables sheet')
                        log('Populating variables from override sheet...')
                        for row in values:
                            SCHEDULE_SPREADSHEET_ID = row[0]
                            TEAMS_PLAYERS_SPREADSHEET_ID = row[1]
                            TEAM_STATS_SPREADSHEET_ID = row[2]
                            TEAMNAMES_RANGE_NAME = row[3]
                            TEAMPLAYERS_RANGE_NAME = row[4]
                            # Conferences no longer exist. Merge Tentatek and Kensa ranges to single variable
                            STATS_RANGE_NAME = row[5]
                            MATCHES_RANGE_NAME = row[6]
                            WEEK_DIFFERENCE = row[7]
                            # Added in Season 5
                            # Variables reference another sheet for matching up team names with their shorthand
                            SHORTHAND_SPREADSHEET_ID = row[8]
                            SHORTHAND_RANGE_NAME = row[9]
        else:
            print('')
            print('Script cannot continue without valid Variables sheet. Quiting...')
            log('User denied to input override info')
            sleep(3)
            clear()
            errexit(2)

    if not values:
        print("Error occured while pulling data from central Variables sheet")
        log('Error occured while pulling data from Variables sheet, asking user for override')
        userconfirm = input('Would you like to create an override that points to another sheet? (Y/n): ')
        if userconfirm.lower() == 'y':
            # TODO: Change override from being another google sheet to instead be a CSV file
            # Likely would create a pre-formatted csv using python rather than downloading a file
            # Maybe I'll get it done before PSL Season 7 ¯\_(ツ)_/¯: lol also no. I would say by season 8 but we all know that won't happen
            valid = False
            while not valid:
                clear()
                tempID = input('What is the ID of the sheet you wish to use: ')
                tempRNG = input('What is the Sheet and Range of the variables (Sheet!A1:Z26): ')
                if tempID != '' and tempRNG != '':
                    clear()
                    print('Testing override sheet...')
                    log('Obtained override info. Testing override validity')
                    values = trysheet(service, tempID, tempRNG, 'Error: Invalid Sheet ID or Range', 2, 'COLUMN')
                        
                    
                    if not values:
                        print('Error: Invalid Sheet ID or Range')
                        log('Error: Override info invalid')
                    else:
                        log('Override info valid')
                        VARIABLES_SPREADSHEET_ID = tempID
                        VARIABLES_RANGE_NAME = tempRNG
                        valid = True
                        if os.path.exists('override.txt'):
                            os.remove('override.txt')
                        open('override.txt','a').close
                        override = open('override.txt','a')
                        override.write(tempID + '\n')
                        override.write(tempRNG)
                        override.close()
                        print('Variables being populated from override variables sheet')
                        log('Populating variables from override sheet...')
                        for row in values:
                            SCHEDULE_SPREADSHEET_ID = row[0]
                            TEAMS_PLAYERS_SPREADSHEET_ID = row[1]
                            TEAM_STATS_SPREADSHEET_ID = row[2]
                            TEAMNAMES_RANGE_NAME = row[3]
                            TEAMPLAYERS_RANGE_NAME = row[4]
                            STATS_RANGE_NAME = row[5]
                            MATCHES_RANGE_NAME = row[6]
                            WEEK_DIFFERENCE = row[7]
                            SHORTHAND_SPREADSHEET_ID = row[8]
                            SHORTHAND_RANGE_NAME = row[9]
    else:
        if overridebool:
            log('Variables being populated from override variables sheet')
            print("Populating variables from override sheet...")
        else:
            log('Variables being populated from central variables sheet')
            print("Populating variables from central sheet...")

        for row in values:
            SCHEDULE_SPREADSHEET_ID = row[0]
            TEAMS_PLAYERS_SPREADSHEET_ID = row[1]
            TEAM_STATS_SPREADSHEET_ID = row[2]
            TEAMNAMES_RANGE_NAME = row[3]
            TEAMPLAYERS_RANGE_NAME = row[4]
            STATS_RANGE_NAME = row[5]
            MATCHES_RANGE_NAME = row[6]
            WEEK_DIFFERENCE = row[7]
            SHORTHAND_SPREADSHEET_ID = row[8]
            SHORTHAND_RANGE_NAME = row[9]

        
    Team1Name = None
    Week = None
    # Ensure Week and Team1Name are valid, loop if not
    while not Week:
        while not Team1Name:
            
            clear()
            # Call Sheets API, grab team names from PSL Schedule
            values = trysheet(service, SCHEDULE_SPREADSHEET_ID, TEAMNAMES_RANGE_NAME, 'Error: Failed to grab team names from PSL Schedule', 4)
            
            
            # Write team 1 selection screen
            log('Obtaining current teams list')
            print('Teams:')
            print('')
            for row in values:
                # Copy to new variable to use later
                Teams = row.copy()
            
            # Print team list for testing purposes
            teamCount = 0
            for team in Teams:
                teamCount += 1
                print(str(teamCount) + '. ' + team)
                log(team + ' added successfully')
            print('')
            
            # Get team selection via number input
            teamSelection = -1
            try:
                teamSelection = int(input('Select team 1: '))
                if teamSelection < 0 or teamSelection > teamCount:
                    clear()
                    teamSelection = -1
                    print('Invalid menu option given, please try again')
                    input()
                else:
                    Team1Name = Teams[teamSelection-1]
            except:
                print('Invalid input given, please try again')
                teamSelection = -1
                input()
            clear()

        log('Variable Team1Name set as ' + Team1Name)

        # Get week number, used both for stream text and to find the opposing team in this matchup  
        Week = input("What week is this match? (number only): ")
        log('Week ' + str(Week) + ' was input')
        clear()

        if Teams:
            # Call Sheets API, grab match schedule
            values = trysheet(service, SCHEDULE_SPREADSHEET_ID, MATCHES_RANGE_NAME, 'Error: Could not pull info from PSL Schedule sheet', 6, 'COLUMNS')

            Team2Name = None
            for row in values:
                # Check to see if selected column matches Team1Name (ignores case, periods, 'the', spaces, and a couple special chars. Seriously y'all I can't filter much more than this to try to prevent issues with names not matching across sheets each season)
                if Team1Name.lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', '') in row[0].lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', ''):
                    # Use the week number given plus an offset to find the match for that team and week
                    log('Trying to pull opposing from PSL Schedule based on Team1Name and Week')
                    Team2Name = row[int(Week)+ int(WEEK_DIFFERENCE)]
                    # Check to ensure the match hasn't already been played (non-0 stats) while it is not a BYE week
                    if not re.search('\\[0+ pts, 0-0-0\\]', Team2Name) and not 'BYE' in Team2Name and not 'TBD' in Team2Name and not Team2Name == '':
                        clear()
                        log('Error occured, week and team cobination shows the match was already played')
                        print("Error: Invalid Team and Week Combination. Team has already played that week's match")
                        print(Team2Name)
                        Team1Name = None
                        Week = None
                        print("Press enter to try again")
                        input()
                        break
                    # Make sure it's not the team's BYE week.
                    elif 'BYE' in Team2Name:
                        clear()
                        log('Error occured, week and team combination shows it is' + Team1Name + "'s BYE week")
                        print("Error: Invalid Team and Week Combination.")
                        print("It is " + Team1Name + "'s BYE week.")
                        Team1Name = None
                        Week = None
                        print("Press Enter to try again")
                        input()
                        break
                    elif 'TBD' in Team2Name:
                        clear()
                        log('Error occured, week and team combination shows this matchup is not determined yet')
                        print('Error: Matchup not yet determined.')
                        Team1Name = None
                        Week = None
                        print("Press Enter to try again")
                        input()
                        break
                    elif Team2Name == '':
                        clear()
                        log('Error occured, cell was empty, likely Bye week')
                        print('Error: Data from this week is empty. Is this a Bye week?')
                        Team1Name = None
                        Week = None
                        print("Press Enter to try again")
                        input()
                        break
                    else:
                        # Remove extra text to get just the team name
                        Team2Name = re.sub("\\[0+ pts, 0-0-0\\] vs ", "", Team2Name)
                        # We found what we need, break to avoid running through the rest of loop
                        break

    # If Team2Name is empty at this point, something broke. Shut EVERYTHING down.
    if not Team2Name:
        print('Error occured while attempting to find Team 2 for the match')
        log('Error occured while attempting to find Team 2 for the match')
        errexit(7)
    log('Variable Team2Name set as ' + Team2Name)
    # Call Sheets API, grab team names and associated players
    values = trysheet(service, TEAMS_PLAYERS_SPREADSHEET_ID, TEAMPLAYERS_RANGE_NAME, 'Error: Could not pull data from Player/Information Sheet', 8)

    # Create empty arrays to append player names to
    Team1Players = []
    Team2Players = []
    log('Attempting to grab players for each team')
    for column in values:
        if Team1Name.lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', '') in column[0].lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', ''):
            y = 0
            for player in column:
                if not y == 0:
                    # Clean up player name, add to team array
                    player = player.replace('*', '')
                    player = player.replace('[V]', '')
                    player = player.replace('[★]', '')
                    player = player.replace('[Ω]', '')
                    player = re.sub(r"\(.*\)", "", player)
                    player = re.sub(r"\[T.*\]", "", player)
                    player = player.rstrip()
                    Team1Players.append(player);
                else:
                    y += 1
        elif Team2Name.lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', '') in column[0].lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', ''):
            y = 0
            for player in column:
                if not y == 0:
                    # Clean up player name, add to team array
                    player = player.replace('*', '')
                    player = player.replace('[V]', '')
                    player = player.replace('[★]', '')
                    player = player.replace('[Ω]', '')
                    player = re.sub(r"\(.*\)", "", player)
                    player = re.sub(r"\[T.*\]", "", player)
                    player = player.rstrip()
                    Team2Players.append(player);
                else:
                    y += 1

    
    if not Team1Players:
        print('Error occured while attempting to get players for ' + Team1Name)
        log('Error occured while attempting to get players for ' + Team1Name)
        input('Press enter to exit')
        errexit(8)
        
    if not Team2Players:
        print('Error occured while attempting to get players for ' + Team2Name)
        log('Error occured while attempting to get players for ' + Team2Name)
        input('Press enter to exit')
        errexit(8)
    
    # Print info for matchup (what the teams are)
    print('This Match:')
    print(Team1Name + " vs " + Team2Name)
    print('')

    # Print Team 1 player list
    print(Team1Name + ' Players:')
    for player in Team1Players:
        log(player + ' successfully added to ' + Team1Name)
        print(player)
    print('')

    # Print Team 2 player list
    print(Team2Name + ' Players:')
    for player in Team2Players:
        log(player + ' successfully added to ' + Team2Name)
        print(player)
    print('')

    if Team1Players and Team2Players:
        # Call Sheets API, grab Tentatek conference first
        values = trysheet(service, TEAM_STATS_SPREADSHEET_ID, STATS_RANGE_NAME, "Error: could not grab team stats", 9)
        Team1Stats = False
        Team2Stats = False
        
        log('Looking for teams in stat sheet')
        for column in values:
            # Check if current row is the team we are looking for (ignores case and space)
            if Team1Name.replace(' ', '').replace('.', '').lower() in column[1].replace(' ', '').replace('.', '').lower():
                # Grab important stats
                log(Team1Name + ' found on stat sheet. Grabbing stats')
                Team1PTS = column[3]
                Team1Wins = column[4]
                Team1Loss = column[6]
                Team1OT = column[7]
                Team1KO = column[8]
                Team1Rank = column[0]
                Team1Stats = True
                if Team2Stats:
                    break
            elif Team2Name.replace(' ', '').replace('.', '').lower() in column[1].replace(' ', '').replace('.', '').lower():
                # Grab important stats
                log(Team2Name + ' found on stat sheet. Grabbing stats')
                Team2PTS = column[3]
                Team2Wins = column[4]
                Team2Loss = column[6]
                Team2OT = column[7]
                Team2KO = column[8]
                Team2Rank = column[0]
                Team2Stats = True
                if Team1Stats:
                    break
        
        if not Team1Stats or not Team2Stats:
            # Stats were not pulled, abort
            if not Team1Stats:
                print('Error: Team 1 stats not found!')
                print('Quitting...')
                sleep(5)
                errexit(13)
            if not Team2Stats:
                print('Error: Team 2 stats not found!')
                print(' Quitting...')
                sleep(5)
                errexit(13)
        
        # Compile all Team 1 Stats to single string, print that string
        log('Combining all ' + Team1Name + "'s stats into a single variable")
        Team1Stats = Team1PTS + " pts | " + Team1Wins + "-" + Team1Loss + "-" + Team1OT
        print(Team1Name + ' Stats:')
        print(Team1Stats)
        print('Rank: ' + str(Team1Rank))
        print("KO's: " + str(Team1KO))
        print('')
        
        # Compile all Team 1 Stats to single string, print that string
        log('Combining all ' + Team2Name + "'s stats into a single variable")
        Team2Stats = Team2PTS + " pts | " + Team2Wins + "-" + Team2Loss + "-" + Team2OT
        print(Team2Name + ' Stats:')
        print(Team2Stats)
        print('Rank: ' + str(Team2Rank))
        print("KO's: " + str(Team2KO))
        print('')

    # Prepare file structure (clear certain old files, create/recreate folders)
    if os.path.exists('Text\Team 1 Members'):
        log('Clearing previous Team 1 Members text files')
        shutil.rmtree('Text\Team 1 Members')

    if os.path.exists('Text\Team 2 Members'):
        log('Clearing previous Team 2 Members text files')
        shutil.rmtree('Text\Team 2 Members')

    if not os.path.exists('Text'):
        os.mkdir('Text')

    if not os.path.exists('Text\Team 1 Members'):
        os.mkdir('Text\Team 1 Members')

    if not os.path.exists('Text\Team 2 Members'):
        os.mkdir('Text\Team 2 Members')

    # Set current path for file manipulation
    dir_path = os.path.dirname(os.path.realpath(__file__))
    
    
    # Set names for logo images to search for
    Team1Logo = Team1Name + '.png'
    Team2Logo = Team2Name + '.png'
    
    # TODO: Set up to automatically download missing logos
    if os.path.exists(dir_path + '\\Team Logos Standardized'):
        # Look through standardized logo images, if match is found copy to Text folder for use on stream
        # Placeholder will be used if no team logo is found
        for root, dirs, files in os.walk(dir_path + '\\Team Logos Standardized'):
            if Team1Logo in files:
                shutil.copyfile(dir_path + '\Team Logos Standardized\\' + Team1Logo, dir_path + '\Text\Team1Logo.png')
            else:
                shutil.copyfile(dir_path + '\Team Logos Standardized\LOGO_NOT_AVAILABLE.png', dir_path + '\Text\Team1Logo.png')
                
            if Team2Logo in files:
                shutil.copyfile(dir_path + '\Team Logos Standardized\\' + Team2Logo, dir_path + '\Text\Team2Logo.png')
            else:
                shutil.copyfile(dir_path + '\Team Logos Standardized\LOGO_NOT_AVAILABLE.png', dir_path + '\Text\Team2Logo.png')
            
        
        # Clear previous scoreboard images
        # This prevents issues on an occasion where a scoreboard icon is not found since there isn't a placeholder like the logos
        if os.path.exists(dir_path + '\Text\Team1Scoreboard.png'):
            os.remove(dir_path + '\Text\Team1Scoreboard.png')
        if os.path.exists(dir_path + '\Text\Team2Scoreboard.png'):
            os.remove(dir_path + '\Text\Team2Scoreboard.png')
    else:
        log('!Folder "Team Logos Standardized" not found, unable to do team logos!')
        print('Error: Folder "Team Logos Standardized" not found, skipping team logo updates!')
    
    Team1Short = None
    Team2Short = None
    values = trysheet(service, SHORTHAND_SPREADSHEET_ID, SHORTHAND_RANGE_NAME, "Error: could not fetch teamname shorthand data", 10)
    for column in values:
        if Team1Name.lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', '') in column[0].lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', ''):
            Team1Short = column[1]
        elif Team2Name.lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', '') in column[0].lower().replace('the', '').replace('.', '').replace(' ', '').replace('#', '').replace('/', ''):
            Team2Short = column[1]
    
    if not Team1Short == None and not Team2Short == None:
        # User shorthand that was pulled to determine file name we are looking for
        Team1Scoreboard = 'PSL SB3 ' + Team1Short + '.png'
        Team2Scoreboard = 'PSL SB3 ' + Team2Short + '.png'
        # Search folder for our files, copy to Text if match is found
        if os.path.exists(dir_path + '\PSL _SB3_ Scoreboard Overlay'):
            for root, dirs, files in os.walk(dir_path + '\\PSL _SB3_ Scoreboard Overlay'):
                if Team1Scoreboard in files:
                    shutil.copyfile(dir_path + '\PSL _SB3_ Scoreboard Overlay\\' + Team1Scoreboard, dir_path + '\Text\Team1Scoreboard.png')
                if Team2Scoreboard in files:
                    shutil.copyfile(dir_path + '\PSL _SB3_ Scoreboard Overlay\\' + Team2Scoreboard, dir_path + '\Text\Team2Scoreboard.png')
        else:
            log('!Folder "PSL _SB3_ Scoreboard Overlay" not found!')
            print('Folder "PSL _SB3_ Scoreboard Overlay" not found. Skipping scoreboard update!')
    else:
        log('!Shorthand name for one of the teams is missing. Skipping scoreboard update!')
        print('Error: Shorthand name for one of the teams is missing. Skipping scoreboard update!')
    
    # Output collected information to text files for use in OBS on stream
    
    # Shorthand text files are fallback for stream in case scoreboard image is not found/doesn't exist 
    log('Writing Team1Shorthand')
    outfile('Text\Team1Shorthand.txt', Team1Short)
    
    log('Writing Team2Shorthand')
    outfile('Text\Team2Shorthand.txt', Team2Short)
    
    log('Writing Team 1 Name')
    outfile('Text\Team 1 Name.txt', Team1Name)

    log('Writing Team 2 Name')
    outfile('Text\Team 2 Name.txt', Team2Name)

    log('Writing Team 1 Stats')
    outfile('Text\Team 1 Stats.txt', Team1Stats)

    log('Writing Team 2 Stats')
    outfile('Text\Team 2 Stats.txt', Team2Stats)

    log('Writing Week')
    outfile('Text\Week.txt', 'Week ' + str(Week))
    
    log('Writing Team 1 Rank')
    outfile('Text\Team 1 Rank.txt', Team1Rank)
    
    log('Writing Team 2 Rank')
    outfile('Text\Team 2 Rank.txt', Team2Rank)
    
    log("Writing Team 1 KO's")
    outfile('Text\Team 1 KO.txt', Team1KO)
    
    log("Writing Team 2 KO's")
    outfile('Text\Team 2 KO.txt', Team2KO)
    
    log('Writing Team 1 Players')
    MemCount = 1
    for player in Team1Players:
        filename = 'Text\Team 1 Members\Member' + str(MemCount) + '.txt'
        filegood = outfile(filename, player)
        
        if filegood:
            log('Player ' + str(MemCount) + ' successful')
        MemCount += 1

    log('Writing Team 2 Players')
    MemCount = 1
    for player in Team2Players:
        filename = 'Text\Team 2 Members\Member' + str(MemCount) + '.txt'
        filegood = outfile(filename, player)
        
        if filegood:
            log('Player ' + str(MemCount) + ' successful')
        MemCount += 1

    print('Update complete! All text on stream should be updated!')
    print('Press Enter to finish.')
    log('Script completed successfully')
    log('~~~~~Script End~~~~~', True)
    input()

if __name__ == '__main__':
    main()