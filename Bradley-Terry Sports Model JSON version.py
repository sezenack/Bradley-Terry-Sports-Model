import xlrd
import openpyxl
import json

'''
The team's rating is calculated from the given parameters
Parameters: record is a tuple (wins, losses)
wfsum is an int, the sum of the weighting factors of the team's opponents
sos_sum is an int, the sum of the strength of schedule of each game of a team
Returns: the team's rating
'''
def RatingCalculator(record, wfsum, sos_sum):
    # Calculates the rating for a team
    if record[1] != 0 and record[0] != 0:
        # if both wins and losses are not 0, ratio can be calculated as normal
        ratio = record[0] / record[1]
    elif record[1] == 0:
        # 0 losses means an infinite win ratio; set it to 25 instead
        ratio = 25
    else:
        # 0 wins means a 0 win ratio; set it to 0.04 instead
        ratio = 1/25
    # Strength of schedule = sum of sos of each game / sum of the weighting factors
    sos = sos_sum / wfsum
    # Each team's rating is their win ratio multiplied by their strength of schedule
    return ratio * sos

# Open json file with utf-8 encoding; insert your own json filename
with open("CBU Post-CT.json", encoding = 'utf-8') as json_file:  
    data = json.load(json_file)

# Initialize list of games (will be tuple of (winning team, losing team))
games = []
# Initalize dictionary of teams mapping their id to their name
teams = dict()
# Dictionary of every team's rating; key = team name, value = rating
ratings = dict()
# Dictionary of every team's record; key = team name, value = list of [wins, losses]
record = dict()
# Dictionary of every team's weighting factor sum; key = team name, value = sum of weighting factors from games
weightingfactorsum = dict()
# Dictionary of every team's strength of schedule sum; key = team name, value = sum of strength of schedule from games
sos_sum = dict()
# Dictionary of every team's new rating calculated; key = team name, value = new rating
new_ratings = dict()
# Dictionary of every team's info; key = team name, value = list of [win %, win ratio, strength of schedule]
info = dict()
# Value to compare every rating to and then do recursion
DELTA = 0.0001
# Main boolean used with while loop for recursion based on flag
done = False
# If all values are within DELTA (i.e. should recursion finish)
flag = True
# Extract information that we need from the data
for t in data["teams"]:
    # teams[id] = name
    teams[t["tid"]] = t["region"]
    # Initialize record to 0-0
    record[t["region"]] = [0, 0]
    # Always start with 100 as every team's rating (using iteration to solve the recursive problem)
    ratings[t["region"]] = 100
    # Initialize the sums to 0
    sos_sum[t["region"]] = 0
    weightingfactorsum[t["region"]] = 0    

# Create a Workbook
wb = openpyxl.Workbook()
# Create sheet to output games to so that postseason games can be easily added
out = wb.active
out.title = "Games"
# Row counter
r = 0
# Go through all the games and add tuple of (winning team, losing team) to list of games
# In addition output the games to a sheet so that postseason games can be added with ease
for g in data["games"]:
    winnerid = g["won"]["tid"]
    loserid = g["lost"]["tid"]
    games.append((teams[winnerid], teams[loserid]))
    # Incrememnt record of each team
    record[teams[winnerid]][0] += 1
    record[teams[loserid]][1] += 1
    # If team is undefeated, we'll get a divide by 0 error, so we set the ratio to 25
    # Note: this formula is most effective when there are no undefeated or winless teams
    # and a chain of wins (or ties) can be made from every team to any other team
    if record[teams[winnerid]][1] == 0:
        ratio = 25
    # If the team is winless, everything would automatically be 0, so set the ratio to 1/25
    elif record[teams[winnerid]][0] == 0:
        ratio = 1/25
    # Win ratio = wins / losses
    else:
        ratio = record[teams[winnerid]][0] / record[teams[winnerid]][1]
    # Add the team's win %, win ratio and sos is intialized to 0 (to be summed later)
    teaminfo = [record[teams[winnerid]][0] / (record[teams[winnerid]][0] + record[teams[winnerid]][1]), ratio, 0]
    # Add the list of team info to the info dict
    info[teams[winnerid]] = teaminfo
    # Repeat for losing team
    if record[teams[loserid]][1] == 0:
        ratio = 25
    elif record[teams[loserid]][0] == 0:
        ratio = 1/25
    # Win ratio = wins / losses
    else:
        ratio = record[teams[loserid]][0] / record[teams[loserid]][1]
    # Add the team's win %, win ratio and sos is intialized to 0 (to be summed later)
    teaminfo = [record[teams[loserid]][0] / (record[teams[loserid]][0] + record[teams[loserid]][1]), ratio, 0]
    # Add the list of team info to the info dict
    info[teams[loserid]] = teaminfo
    if r == 0:
        out.cell(row = r + 1, column = 1).value = 'Winning Team'
        out.cell(row = r + 1, column = 2).value = 'Winning Score'
        out.cell(row = r + 1, column = 3).value = 'Losing Team'
        out.cell(row = r + 1, column = 4).value = 'Losing Score'
    else:
        out.cell(row = r + 1, column = 1).value = teams[winnerid]
        out.cell(row = r + 1, column = 2).value = g["won"]["pts"]
        out.cell(row = r + 1, column = 3).value = teams[loserid]
        out.cell(row = r + 1, column = 4).value = g["lost"]["pts"]
    r += 1

# Recursion for getting accurate ratings
while not done:
    # Initialize the flag to True each time
    flag = True
    
    # For every game, calculate:
    for game in games:
        # Weighting factor (1 divided by the sum of the ratings of the 2 teams)
        wf = 1 / (ratings[game[0]] + ratings[game[1]])
        # Add the weighting factor to the dict for each team
        weightingfactorsum[game[0]] += wf
        weightingfactorsum[game[1]] += wf
        # SOS for each game is the opponent's rating multiplied by the weighting factor
        # Add the SOS for the game to the SOS sum for each team
        sos_sum[game[0]] += (wf * ratings[game[1]])
        sos_sum[game[1]] += (wf * ratings[game[0]])
    
    # For every team, calculate:
    for key in ratings:
        # New rating for the team equals the team's winning ratio multiplied by their SOS
        # SOS equals the SOS sum divided by the weighting factor sum
        new_ratings[key] = RatingCalculator(record[key], weightingfactorsum[key], sos_sum[key])
        # Update the SOS for the team
        info[key][2] = sos_sum[key] / weightingfactorsum[key]
        # If the difference between old rating and new rating <= DELTA
        # If flag is true, that means so far every team's new rating has been within DELTA
        # (since flag is initialized to true)
        if abs(ratings[key] - new_ratings[key]) <= DELTA and flag:
            done = True
        # If the difference is greater than DELTA, we must continue the recursion
        # If flag is false, one team has already failed the DELTA test and we must continue the recursion
        else:
            flag = False
            done = False
    # After going through all teams, update ratings
    ratings = new_ratings

# Sort the teams by their rating
sortedratings = sorted(ratings.items(), key=lambda kv: kv[1], reverse = True)
# Create new sheet for output
wb.create_sheet('Output')
# For outputting the results into the spreadsheet
out = wb.get_sheet_by_name('Output')
for r in range(len(sortedratings)):
    # Header row
    if r == 0:
        out.cell(row = r + 1, column = 1).value = 'Team'
        out.cell(row = r + 1, column = 2).value = 'Rating'
        out.cell(row = r + 1, column = 3).value = 'Win % Rank'
        out.cell(row = r + 1, column = 4).value = 'Record'
        out.cell(row = r + 1, column = 5).value = 'Win %'
        out.cell(row = r + 1, column = 6).value = 'Win Ratio'
        out.cell(row = r + 1, column = 7).value = 'SOS Rank'
        out.cell(row = r + 1, column = 8).value = 'SOS'
    else:
        # Eight columns
        for col in range(8):
            # Col 0 is the team name, col 1 is the rating
            if col < 2:
                out.cell(row = r + 1, column = col + 1).value = sortedratings[r - 1][col]
            # Formula to calculate the rank of winning %
            elif col == 2:
                out.cell(row = r + 1, column = col + 1).value = "=RANK(E" + str(r + 1) + \
                    ",E2:E" + str(len(sortedratings)) + ")"
            # Output record
            elif col == 3:
                out.cell(row = r + 1, column = col + 1).value = str(record[sortedratings[r - 1][0]][0]) \
                    + "-" + str(record[sortedratings[r - 1][0]][1])
            # Output win % (col 4) and win ratio (col 5)
            elif col < 6:
                out.cell(row = r + 1, column = col + 1).value = info[sortedratings[r - 1][0]][col - 4]
            # Formula to calculate the SOS rank
            elif col == 6:
                out.cell(row = r + 1, column = col + 1).value = "=RANK(H" + str(r + 1) + \
                    ",H2:H" + str(len(sortedratings)) + ")"
            # Output SOS (col 7)
            else:
                out.cell(row = r + 1, column = col + 1).value = info[sortedratings[r - 1][0]][col - 5]

# Save the sheet with the output
wb.save('Bradley-Terry Spreadsheet JSON CBU.xlsx')
