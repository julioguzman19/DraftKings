import pandas as pd
from openpyxl import load_workbook
from pulp import LpMaximize, LpProblem, LpStatus, lpSum, LpVariable
import requests
from bs4 import BeautifulSoup

# # Load your CSV
# df = pd.read_csv('DKSalaries.csv')

# # Concatenate 'Name' and 'TeamAbbrev'
# df['Name'] = df['Name'].astype(str) + ' ' + df['TeamAbbrev'].astype(str)

# # Save as Excel
# df.to_excel('DKSalaries.xlsx', index=False)

# # Load the workbook
# wb = load_workbook('DKSalaries.xlsx')

# # ------------------------
# # Select the active sheet
# sheet = wb['Sheet1'] 

# # Define the columns you want to keep
# columns_to_keep = ['Position', 'Name', 'ID', 'Salary']

# # Get the max column count
# max_column = sheet.max_column

# # Traverse in reverse order (to avoid index change problem during deletion)
# for i in range(max_column, 0, -1):
#     # If the column header is not in columns_to_keep list, delete it
#     if sheet.cell(row=1, column=i).value not in columns_to_keep:
#         sheet.delete_cols(i)

# # Add new column headers
# sheet['E1'] = 'PredictedPts'  # Assuming 'E' is the next empty column

# # Save the workbook
# wb.save('DKSalaries.xlsx')

# # Load the .xlsx file
# df = pd.read_excel('DKSalaries.xlsx', sheet_name='Sheet1')  

# # -----------------------------------------------------
# # Create the model
# model = LpProblem(name="optimal-lineup", sense=LpMaximize)

# # Create decision variables
# player_vars = LpVariable.dicts("player", df.index, cat='Binary')

# # Add objective function to the model
# model += lpSum([df.loc[i, 'PredictedPts'] * player_vars[i] for i in df.index])

# # Add salary constraint
# model += lpSum([df.loc[i, 'Salary'] * player_vars[i] for i in df.index]) <= 50000

# # Add position constraints
# positions_needed = {'QB': 1, 'RB': 2, 'WR': 3, 'TE': 1, 'DST': 1}
# flex_positions = ['RB', 'WR', 'TE']
# for position, count in positions_needed.items():
#     model += lpSum([player_vars[i] for i in df[df['Position'] == position].index]) == count

# # Handle the FLEX position separately
# model += lpSum([player_vars[i] for i in df[df['Position'].isin(flex_positions)].index]) == sum(positions_needed.values()) + 1  # +1 because of FLEX

# # Add constraint so that each player can only be picked at most once
# model += lpSum([player_vars[i] for i in df.index]) <= 10

# # Add constraint for unique player IDs
# selected_ids = set()
# for i in df.index:
#     player_id = df.loc[i, 'ID']
#     if player_id in selected_ids:
#         model += player_vars[i] == 0
#     else:
#         selected_ids.add(player_id)

# # Solve the model
# status = model.solve()

# # Print the status of the solved LP
# print(f"Status: {LpStatus[model.status]}")

# # Get the players in the optimal lineup
# lineup = df.iloc[[i for i in df.index if player_vars[i].value() == 1]]
# lineup = lineup[lineup['Position'].isin(positions_needed)]
# print(lineup)




BASE_URL = "https://www.fantasypros.com/nfl/projections/{}.php?week=1"
POSITIONS = ['qb', 'rb', 'wr', 'te', 'dst']

# A dictionary to map each position to its relevant stats and indices
position_stats_mapping = {
    'qb': {
        'indices': [2, 3, 4, 6, 7],
        'columns': ['PassYards', 'PassTD', 'Interceptions', 'RushYards', 'RushTD']
    },
    'rb': {
        'indices': [1, 2, 3, 4, 5],
        'columns': ['RushYards', 'RushTD', 'Receptions', 'ReceiveYards', 'ReceiveTD']
    },
    'wr': {
        'indices': [0, 1, 2, 4],
        'columns': ['Receptions', 'ReceiveYards', 'ReceiveTD', 'RushYards']
    },
    'te': {
        'indices': [0, 1, 2],
        'columns': ['Receptions', 'ReceiveYards', 'ReceiveTD']
    },
    'dst': {
        'indices': [0, 1,2, 6],
        'columns': ['Sacks', 'Interceptions', 'FumbleRecover' 'PointsAllowed']
    }
}

def scrape_position(position):
    url = BASE_URL.format(position)
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    player_name = soup.select_one('td.player-label a.player-name').text
    stats_elements = soup.select('td.center')

    stats = [float(stats_elements[idx].text) for idx in position_stats_mapping[position]['indices']]
    data = {column: stat for column, stat in zip(position_stats_mapping[position]['columns'], stats)}

    df = pd.DataFrame([data], columns=position_stats_mapping[position]['columns'])
    df.insert(0, "Player", player_name)
    return df

all_data = {}

# Scrape data for each position
for position in POSITIONS:
    all_data[position] = scrape_position(position)

# Example to print QB data
print(all_data['dst'])