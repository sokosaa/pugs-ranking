import pandas as pd
import math
import matplotlib.pyplot as plt
import numpy as np
import os
import subprocess
import platform

class Player:
    def __init__(self, name, tank_sr=2500, damage_sr=2500, support_sr=2500):
        self.name = name
        self.tank_sr = tank_sr
        self.damage_sr = damage_sr
        self.support_sr = support_sr

    def update_sr(self, role, k, actual_outcome, expected_outcome):
        if role == "tank":
            self.tank_sr += k * (actual_outcome - expected_outcome)
        elif role == "damage":
            self.damage_sr += k * (actual_outcome - expected_outcome)
        elif role == "support":
            self.support_sr += k * (actual_outcome - expected_outcome)


def calculate_k(sr): # adjust this as wanted
    return 50
    # if sr < 2000:
    #     return 50
    # elif 2000 <= sr <= 2500:
    #     return 50
    # elif 2500 <= sr <= 3000:
    #     return 50
    # elif 3000 <= sr <= 3500:
    #     return 50
    # elif 3500 <= sr <= 4000:
    #     return 50
    # else:
    #     return 50

    # # Parameters
    # k_min = 16
    # k_max = 32
    # mid_point = 2500
    # steepness = 0.004
    
    # # Logistic function
    # k = k_min + (k_max - k_min) / (1 + math.exp(-steepness * (sr - mid_point)))
    
    # return k

def plot_k():
    ratings = np.arange(0, 5000, 10)
    k_values = [calculate_k(rating) for rating in ratings]

    plt.plot(ratings, k_values)
    plt.xlabel('Player Rating')
    plt.ylabel('K Value')
    plt.title('K Value as a Function of Player Rating')
    plt.grid()
    plt.show()

def expected_outcome(r_a, r_b):
    return 1 / (1 + 10**((r_b - r_a) / 400))  # og elo formula: P = 1 / (1 + 10^((r_b - r_a) / D))  [P = probabiliy of win with 200 rating difference]

# def check_missing_players(team_a, team_b, df):
#     listed_players = df.iloc[0]
#     print(listed_players[9])
#     missing_players = []
#     for player in team_a + team_b:
#         if not player[0] or player[0] == 'empty' or (player[0] and player[0].name not in df['Player'].tolist()):
#             missing_players.append(player[0].name if player[0] else 'unknown')
#     if missing_players:
#         print(f"The following players are missing: {', '.join(missing_players)}. Skipping SR/rank update.")
#         return False
#     return True

# def check_missing_player(i, df):
#     player_name = df.iloc[0, i]
#     if pd.isna(player_name) or player_name not in df['Player'].tolist():
#         return player_name

# def check_missing_players(team_a, team_b, df):
#     missing_players = []
#     for i in range(10):
#         player_name = check_missing_player(i, df)
#         missing_players.append(player_name)
#     if any(missing_players):
#         print(f"The following players are missing: {', '.join(missing_players)}. Skipping SR/rank update.")
#         return False
#     else:
#         return True

def check_missing_player(i, df):
    player_name = df.iloc[0, i]
    if player_name not in df['Player'].tolist():
        return player_name

def check_missing_players(team_a, team_b, df):
    missing_players = [check_missing_player(i, df) for i in range(10)]
    missing_players = [p for p in missing_players if p is not None]
    if missing_players:
        print(f"Missing Players: {', '.join(missing_players)}. Skipping SR update.")
        return False
    else:
        return True



def update_ratings(team_a, team_b, outcome, df):
    # Check that all players exist
    check_missing_players(team_a, team_b, df)
    if not check_missing_players(team_a, team_b, df):
        return False

    # Get the SR for each player based on their role
    sr_a = [getattr(player[0], f"{player[1]}_sr") for player in team_a]
    sr_b = [getattr(player[0], f"{player[1]}_sr") for player in team_b]

    avg_sr_a = sum(sr_a) / len(sr_a)
    avg_sr_b = sum(sr_b) / len(sr_b)

    # Calculate expected outcomes
    expected_a = expected_outcome(avg_sr_a, avg_sr_b)
    expected_b = expected_outcome(avg_sr_b, avg_sr_a)

    # Calculate actual outcomes
    if outcome == "win":
        actual_a, actual_b = 1, 0
    elif outcome == "loss":
        actual_a, actual_b = 0, 1
    else:  # Draw
        actual_a, actual_b = 0.5, 0.5

    # Update ratings for team A
    for player, role in team_a:
        player_sr = getattr(player, f"{role}_sr")
        k = calculate_k(player_sr)
        player.update_sr(role, k, actual_a, expected_a)

    # Update ratings for team B
    for player, role in team_b:
        player_sr = getattr(player, f"{role}_sr")
        k = calculate_k(player_sr)
        player.update_sr(role, k, actual_b, expected_b)
    
    return True


def get_player_by_name(players, name):
    for player in players:
        if player.name == name:
            return player
    return None

def initialize_players(df):
    players = {}
    for index, row in df.iterrows():
        if index == 0:
            continue

        name = row[df.columns[11]]
        if pd.isna(name):
            continue

        tank_sr = row[df.columns[12]]
        damage_sr = row[df.columns[13]]
        support_sr = row[df.columns[14]]
        player = Player(name, tank_sr, damage_sr, support_sr)
        players[name] = player
    return players


def process_single_match(df, players):
    row = df.iloc[0]
    
    team_roles = ["tank", "damage", "damage", "support", "support"]
    team_a = [(players.get(row.iloc[i]), team_roles[i]) for i in range(5)]
    team_b = [(players.get(row.iloc[i]), team_roles[i - 5]) for i in range(5, 10)]

    outcome_code = row.iloc[10]
    if outcome_code == 1:
        outcome = "win"
    elif outcome_code == 2:
        outcome = "loss"
    else:
        outcome = "draw"

    return update_ratings(team_a, team_b, outcome, df)



def save_updated_ratings(file_name, df, players):
    for index, player in enumerate(players.values()):
        df.iat[index + 1, 12] = player.tank_sr
        df.iat[index + 1, 13] = player.damage_sr
        df.iat[index + 1, 14] = player.support_sr

    # Add a new column for the peak SR
    df['High'] = df.iloc[:, 12:15].apply(max, axis=1)

    # Sort by the 'Peak SR' column
    sorted_df = df.loc[1:].sort_values(by='High', ascending=False, ignore_index=True)
    df.iloc[1:, 11:16] = sorted_df.iloc[:, 11:16]

    df.to_excel(file_name, index=False)



def shift_match_data_down(file_name,df):
    # Shift rows down by 1 for columns A-K
    df.iloc[:, 0:11] = df.iloc[:, 0:11].shift(1)
    df.to_excel(file_name, index=False)

def close_excel_file(file_name):
    if platform.system() == 'Windows':
        os.system(f'taskkill /F /IM excel.exe /FI "WindowTitle eq {file_name} - Excel"')
    else:
        print("Closing Excel not supported on this platform.")

def open_excel_file(file_name):
    if platform.system() == 'Windows':
        os.startfile(file_name)
    else:
        print("Opening Excel not supported on this platform.")

def process_data_file(file_name):
    df = pd.read_excel(file_name)

    players = initialize_players(df)
    update_successful = process_single_match(df, players)
    save_updated_ratings(file_name, df, players)
    if update_successful:
        shift_match_data_down(file_name, df)
        print('Ranks Updated Successfully')
    else:
        print("Ranks update failed.")
    return update_successful


if __name__ == "__main__":
    
    file_name = "rat.xlsx"
    close_excel_file(file_name)
    process_data_file(file_name)
    open_excel_file(file_name)
    # plot_k()
