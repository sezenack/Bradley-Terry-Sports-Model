import openpyxl
import json

# Open json file with utf-8 encoding; insert your own json filename
with open("College_Basketball__80_League_295_2057_CT_summary.json", encoding = 'utf-8-sig') as json_file:  
    data = json.load(json_file)

# Initialize list of games (will be tuple of (winning team, losing team))
games = []
# Initalize dictionary of teams mapping their id to their name
teams = dict()
# Dictionary of every team's rating; key = team name, value = rating
ratings = dict()
# Dictionary of every team's record; key = team name, value = list of [wins, losses]
record = dict()
# Dictionary of every team's expected wins; key = team name, value = expected wins
expected = dict()
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

# Create a Workbook
wb = openpyxl.Workbook()
# Create sheet to output games to so that postseason games can be easily added
out = wb.active
out.title = "Games"
# Row counter
r = 1
# First row
out.cell(row = r, column = 1).value = 'Winning Team'
out.cell(row = r, column = 2).value = 'Winning Score'
out.cell(row = r, column = 3).value = 'Losing Team'
out.cell(row = r, column = 4).value = 'Losing Score'
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
    out.cell(row = r + 1, column = 1).value = teams[winnerid]
    out.cell(row = r + 1, column = 2).value = g["won"]["pts"]
    out.cell(row = r + 1, column = 3).value = teams[loserid]
    out.cell(row = r + 1, column = 4).value = g["lost"]["pts"]
    r += 1

# Recursion for getting accurate ratings
while not done:
    # Initialize the flag to True each time
    flag = True
    
    # Clear expected wins each iteration
    expected.clear()
    # For every game, calculate:
    for game in games:
        # Weighting factor (1 divided by the sum of the ratings of the 2 teams)
        wf = 1 / (ratings[game[0]] + ratings[game[1]])
        # Check to see if each team is in expected wins dictionary
        if game[0] in expected:
            # Multiply team's rating by weighting factor and add to sum
            expected[game[0]] += (ratings[game[0]] * wf)
        else:
            # Multiply team's rating by weighting factor and initialize as expected wins
            expected[game[0]] = (ratings[game[0]] * wf)
        if game[1] in expected:
            # Multiply team's rating by weighting factor and add to sum
            expected[game[1]] += (ratings[game[1]] * wf)
        else:
            # Multiply team's rating by weighting factor and initialize as expected wins
            expected[game[1]] = (ratings[game[1]] * wf)
    
    # For every team, calculate:
    for key in ratings:
        # New rating for the team equals the team's wins divided by expected wins multiplied by the old rating
        new_ratings[key] = (record[key][0] / expected[key]) * ratings[key]
        # Update the SOS for the team
        info[key][2] = new_ratings[key] / info[key][1]
        # If the difference between old rating and new rating <= DELTA
        # If flag is true, that means so far every team's new rating has been within DELTA
        # (since flag is initialized to true)
        if abs(ratings[key] - new_ratings[key]) <= DELTA and abs(record[key][0] - expected[key]) <= DELTA and flag:
            done = True
        # If the difference is greater than DELTA, we must continue the recursion
        # If flag is false, one team has already failed the DELTA test and we must continue the recursion
        else:
            flag = False
            done = False
    # After going through all teams, update ratings
    ratings = new_ratings

# Scale the ratings to an average of 100
for i in range(10):
    scale_wins = 0
    for key in ratings:
        scale_wins += 100 / (100 + ratings[key])
    # The denominator should be the number of teams in the league divided by 2
    scale = scale_wins / 50
    
    # Adjust every team's rating and SOS according to scale
    for key in ratings:
        ratings[key] *= scale
        info[key][2] = ratings[key] / info[key][1]

# Sort the teams by their rating
sortedratings = sorted(ratings.items(), key=lambda kv: kv[1], reverse = True)
# Create new sheet for output
wb.create_sheet('Output')
# For outputting the results into the spreadsheet
out = wb.get_sheet_by_name('Output')
for r in range(len(sortedratings) + 1):
    # Header row
    if r == 0:
        out.cell(row = r + 1, column = 1).value = 'Rank'
        out.cell(row = r + 1, column = 2).value = 'Team'
        out.cell(row = r + 1, column = 3).value = 'Rating'
        out.cell(row = r + 1, column = 4).value = 'Win % Rank'
        out.cell(row = r + 1, column = 5).value = 'Record'
        out.cell(row = r + 1, column = 6).value = 'Win %'
        out.cell(row = r + 1, column = 7).value = 'Win Ratio'
        out.cell(row = r + 1, column = 8).value = 'SOS Rank'
        out.cell(row = r + 1, column = 9).value = 'SOS'
    else:
        # Nine columns
        for col in range(9):
            # Col 0 is the rank
            if col == 0:
                out.cell(row = r + 1, column = col + 1).value = r
            # Col 1 is the team name, col 2 is the rating
            elif col < 3:
                out.cell(row = r + 1, column = col + 1).value = sortedratings[r - 1][col - 1]
            # Formula to calculate the rank of winning %
            elif col == 3:
                out.cell(row = r + 1, column = col + 1).value = "=RANK(F" + str(r + 1) + \
                    ",F2:F" + str(len(sortedratings) + 1) + ")"
            # Output record
            elif col == 4:
                out.cell(row = r + 1, column = col + 1).value = str(record[sortedratings[r - 1][0]][0]) \
                    + "-" + str(record[sortedratings[r - 1][0]][1])
            # Output win % (col 5) and win ratio (col 6)
            elif col < 7:
                out.cell(row = r + 1, column = col + 1).value = info[sortedratings[r - 1][0]][col - 5]
            # Formula to calculate the SOS rank
            elif col == 7:
                out.cell(row = r + 1, column = col + 1).value = "=RANK(I" + str(r + 1) + \
                    ",I2:I" + str(len(sortedratings) + 1) + ")"
            # Output SOS (col 8)
            else:
                out.cell(row = r + 1, column = col + 1).value = info[sortedratings[r - 1][0]][col - 6]

# Save the sheet with the output
wb.save('Bradley-Terry Spreadsheet JSON NCBCA.xlsx')
