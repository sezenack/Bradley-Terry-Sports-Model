# Bradley-Terry-Sports-Model
Calculates Bradley-Terry ratings given teams, games and records

Getting Started: \
You will need to download xlrd and openpyxl for python \
For V1, you need a xlsx file with 2 sheets: \
Sheet 1 (Team Info) - Column A (Team Name), Column B (Wins), Column C (Losses), Column D (Ties, optional) and any other stats you may want to look at \
Sheet 2 (Every game played) - Column A (Winning Team), Column B (Winning Team Score), Column C (Losing Team), Column D (Losing Team Score) \
Note that for sheet 1, only the first 3 columns are needed. For sheet 2, scores are optional, but there still must be a placeholder column in B \
I have included an example spreadsheet \
For V2, you need a json file that has all the info requested when the program reads the file

The Bradley-Terry System is based on logistic regression essentially meaning that the ratings are calculated directly from win-loss results by all teams.  The strength of schedule is calculated directly from the ratings themselves, meaning that the ratings cannot be easily distorted by teams who rack up the wins against weak opposition.  This is what makes it different from many other ranking systems because that is not case with most systems.  Another cool thing about this rating system is that the ratings are done on an odds basis.  If Team A has a rating that's double Team B's rating, they would be expected to win twice as many games versus Team B than they lose.  Essentially, Team A would be expected to win 2/3 of the time that they play Team B. Pretty cool right?  This model can be used for predictions and odds with this, however, it should be noted that margin of victory (and defeat), injuries and hot (and cold) streaks do not factor into the ratings whatsoever.  You are exactly as good as your record and the opponents you played.  Anything else introduces subjectivity anyways, but that's important to remember before trying to use the ratings.

The ratings are calculating by multiplying each team's winning ratio by their strength of schedule.  A team's strength of schedule is the weighted average of their opponents' ratings.  For each team and each of their games, there is a weighting factor which is 1 divided by the sum of the team's rating and the their opponent's rating.  Multiply this weighting factor by their opponent's rating.  Sum up all these products from all their games, and divide that sum by the sum of the weighting factors and that is the team's strength of schedule.  This is a recursive algorithm, which is solved by iteration.  Start out by giving every team a rating of 100 (the average rating in this system), and stop recursively calculating the ratings when there is very little difference between the new ratings and old ratings.  If you want more information, this rating system is already applied to College Hockey (called KRACH) and College Hockey News does an excellent job explaining the rating system https://www.collegehockeynews.com/info/?d=krach

Please note that while this is an awesome rating system, it is most effective when a chain of wins (and/or ties) can be made from every team to any other team.  This means that a decently large sample size of games is needed, and undefeated teams/winless teams should not exist.  However, ratings may still be calculated without this; they just won't be as useful/relevant.
