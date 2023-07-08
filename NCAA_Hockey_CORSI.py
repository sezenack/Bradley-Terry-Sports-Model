import json
import gspread
import urllib.request, urllib.error, urllib.parse
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

STARTDATE = datetime.now()
FSTARTDATE = STARTDATE.strftime("%Y%m%d")

def writeOutput(teamstats, filename):
    # Make ratings dict for output
    ratings = dict()
    for t in teamstats:
        ratings[t] = teamstats[t]["rating"]

    # Sort the teams by their rating
    sortedratings = sorted(ratings.items(), key=lambda kv: kv[1], reverse = True)

    # Create account object for google sheet writing
    gc = gspread.oauth()
    # For outputting the results into the spreadsheet
    wb = gc.open(filename)
    out = wb.sheet1
    
    # Array for entire sheet
    sheet = []
    for r in range(len(sortedratings) + 2):
        # Row array
        row = []
        # Timestamp for viewers to know most recent update
        if r == 0:
            row = ['Last Updated:', datetime.now().strftime("%m/%d/%Y, %H:%M:%S")]
        # Header row
        elif r == 1:
            row = ['Rank','Team','Rating','Avg Corsi % Rank','Avg Corsi %','Corsi Ratio','SOS Rank','SOS']
        else:
            # Nine columns
            for col in range(8):
                # Col 0 is the rank
                if col == 0:
                    row.append(r - 1)
                # Col 1 is the team name, col 2 is the rating
                elif col < 3:
                    row.append(sortedratings[r - 2][col - 1])
                # Formula to calculate the rank of avg corsi %
                elif col == 3:
                    row.append("=RANK(E" + str(r + 1) + ",E3:E" + str(len(sortedratings) + 2) + ")")
                # Output avg corsi % (col 4) and avg corsi ratio (col 5)
                elif col < 6:
                    row.append(teamstats[sortedratings[r - 2][0]]["info"][col - 4])
                # Formula to calculate the SOS rank
                elif col == 6:
                    row.append("=RANK(H" + str(r + 1) + ",H3:H" + str(len(sortedratings) + 2) + ")")
                # Output SOS (col 7)
                else:
                    row.append(teamstats[sortedratings[r - 2][0]]["info"][col - 5])

        # Add row to sheet
        sheet.append(row)

    # Add output
    out.update("A1:H" + str(len(sheet)), sheet, raw=False)

def scaleRatings(teamstats):
    # Scale the ratings to an average of 50
    for _ in range(10):
        scale_wins = 0
        for key in teamstats:
            scale_wins += 50 / (50 + teamstats[key]["rating"])
        # Divide by half the total number of teams
        scale = scale_wins / (len(teamstats) / 2)
        
        # Adjust every team's rating and SOS according to scale
        for key in teamstats:
            teamstats[key]["rating"] *= scale
            teamstats[key]["info"][2] = teamstats[key]["rating"] / teamstats[key]["info"][1]

def calculateRatings(teamstats, games):
    # Value to compare every rating to and then do recursion
    DELTA = 0.0001
    # Main boolean used with while loop for recursion based on flag
    DONE = False
    # If all values are within DELTA (i.e. should recursion finish)
    FLAG = True

    # Recursion for getting accurate ratings
    while not DONE:
        # Initialize the flag to True each time
        FLAG = True
        
        # Clear expected wins each iteration
        for t in teamstats:
            teamstats[t]["expected"] = 0
        
        # For every game, calculate:
        for game in games:
            # Weighting factor (1 divided by the sum of the ratings of the 2 teams)
            wf = 1 / (teamstats[game[0]]["rating"] + teamstats[game[1]]["rating"])
            # Multiply team's rating by weighting factor and add to sum
            teamstats[game[0]]["expected"] += (teamstats[game[0]]["rating"] * wf)
            # Multiply team's rating by weighting factor and add to sum
            teamstats[game[1]]["expected"] += (teamstats[game[1]]["rating"] * wf)
        
        # For every team, calculate:
        for key in teamstats:
            # New rating for the team equals the team's wins divided by expected wins multiplied by the old rating
            wins = teamstats[key]["record"][0]
            teamstats[key]["new_rating"] = (wins / teamstats[key]["expected"]) * teamstats[key]["rating"]
            
            # Update the SOS for the team
            teamstats[key]["info"][2] = teamstats[key]["new_rating"] / teamstats[key]["info"][1]
            
            # If the difference between old rating and new rating <= DELTA
            # If flag is true, that means so far every team's new rating has been within DELTA
            # (since flag is initialized to true)
            if abs(teamstats[key]["rating"] - teamstats[key]["new_rating"]) <= DELTA \
            and abs(wins - teamstats[key]["expected"]) <= DELTA and FLAG:
                DONE = True
            # If the difference is greater than DELTA, we must continue the recursion
            # If flag is false, one team has already failed the DELTA test and we must continue the recursion
            else:
                FLAG = False
                DONE = False
        
        # After going through all teams, update ratings
        for t in teamstats:
            teamstats[t]["rating"] = teamstats[t]["new_rating"]

def updateInfo(teamstats):
    # Update info after every game is read
    for key in teamstats:
        # If team is undefeated or winless, we'll get a divide by 0 error
        # Note: this formula is most effective when there are no undefeated or winless teams
        # and a chain of wins (or ties) can be made from every team to any other team
        wins = teamstats[key]["record"][0]
        ratio = wins / teamstats[key]["record"][1]
        # Add the team's avg corsi %, avg corsi ratio and sos is intialized to 0 (to be summed later)
        teaminfo = [wins / (teamstats[key]["record"][0] + teamstats[key]["record"][1]) * 100, ratio, 0]
        # Add the list of team info to the info dict
        teamstats[key]["info"] = teaminfo

def getCorsiData(filename):
    # List of all games with links to corsi data
    # Games get cached in json so we only read current day
    url = "https://www.collegehockeynews.com/schedules/?date={}&de=20230410".format(FSTARTDATE)
    f = urllib.request.urlopen(url)
    html = f.read()
    f.close()
    with open(filename, "r") as file:
        gameDict = json.load(file)

    # Get table with links to set up iteration for scraping
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table",{"class":"data schedule full"})
    table = table.find("tbody")   
    metricLinks = table.find_all("td", {"class": "noprint lomobile center m"})

    for metric in metricLinks:
        # Open metrics page
        if metric.get_text().strip() != "":
            # Set up scraping of metrics
            url = "https://www.collegehockeynews.com" + metric.find("a")["href"]
            f = urllib.request.urlopen(url)
            html = f.read()
            f.close()
            
            # Scrape team names, date of game and close corsi from page
            try:
                soupGame = BeautifulSoup(html, "html.parser")
                aTeam, hTeam = soupGame.find('h2').get_text().split(" vs. ")
                date = datetime.strptime(soupGame.find('div', {'id': 'content'}).find('h3').get_text(), "%A, %B %d, %Y").date()
                shotTables = soupGame.find_all("table")
                aCF = int(shotTables[0].find("tfoot").find("td",{"class":"cls tsa"}).get_text())
                hCF = int(shotTables[1].find("tfoot").find("td",{"class":"cls tsa"}).get_text())
            except Exception as e:
                print(e)
                continue

            # Create a unique key for each game and add info to the dictionary
            key = str(date) + aTeam + hTeam
            print(key)
            gameDict[key] = {"date": str(date), "aTeam": aTeam, "aCF": aCF,  "hTeam": hTeam, "hCF": hCF, "url": url}

    # Dump updated data into json as cache
    with open(filename, "w") as file:
        json.dump(gameDict, file)

    return gameDict

def readGames(teamstats, games, filename):
    # Get corsi data from local json caching results using function
    data = getCorsiData(filename)

    # Read in every game with the teams
    for game in data:
        info = data[game]
        team1 = info["aTeam"]
        team1closecorsi = info["aCF"]
        team2 = info["hTeam"]
        team2closecorsi = info["hCF"]
        
        # Initialize if necessary
        if team1 not in teamstats:
            # Initialize team dict
            teamstats[team1] = dict()
            # Initialize record to 0-0
            teamstats[team1]["record"] = [0, 0]
            # Always start with 50 as every team's rating (using iteration to solve the recursive problem)
            teamstats[team1]["rating"] = 50
        if team2 not in teamstats:
            # Initialize team dict
            teamstats[team2] = dict()
            # Initialize record to 0-0
            teamstats[team2]["record"] = [0, 0]
            # Always start with 50 as every team's rating (using iteration to solve the recursive problem)
            teamstats[team2]["rating"] = 50
        
        # Total shot attempts
        total = team1closecorsi + team2closecorsi
        # Use decimal wins
        teamstats[team1]["record"][0] += team1closecorsi / total
        teamstats[team2]["record"][1] += team1closecorsi / total
        teamstats[team1]["record"][1] += team2closecorsi / total
        teamstats[team2]["record"][0] += team2closecorsi / total
        
        games.append((team1, team2))

def main(read):
    # List of tuples which are the games; each tuple is (winner, loser)
    games = []

    # Mega dictionary with all info
    # key = team name
    # value = dictionary with: rating (int), record (list of [wins, losses]), expected wins (int)
    # new rating (int) and info (list of [avg corsi %, corsi ratio, strength of schedule])
    teamstats = dict()

    readGames(teamstats, games, read)
    updateInfo(teamstats)
    calculateRatings(teamstats, games)
    scaleRatings(teamstats)

    return teamstats

if __name__ == "__main__":
    # Replace with your filenames
    teamstats = main('../closeCorsi.json')
    writeOutput(teamstats, "SNACC (Stephen's New Adjusted Close Corsi)")
