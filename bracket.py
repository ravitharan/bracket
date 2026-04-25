#! /usr/bin/env python3

import sys
import openpyxl
from itertools import combinations
import random

ATTENDANCE_SHEET = "attendance"

SCHEDULE_SHEET = "quali_schedule"
WINNER_POINTS = 1

SCORE_SHEET = "quali_score"

LEFT_SCORE_COL = 'C'
LEFT_POINTS_COL = 'D'
RIGHT_POINTS_COL = 'F'
RIGHT_SCORE_COL = 'G'

COURTS_COUNT = 5
MENS_COURT_COUNT = 3
CHILD_COURT_INX = MENS_COURT_COUNT
CRT5_COURT_INX = COURTS_COUNT - 1

def retrive_teams(wb):
    ws = wb[ATTENDANCE_SHEET]
    teams = {}
    groups = [[] for i in range(6)]
    quali_matches = []
    team_ranks = []

    for row in range(2, ws.max_row+1):
        rank = ws.cell(row, 1).value
        team = ws.cell(row, 2).value
        player1 = ws.cell(row, 3).value
        attend1 = ws.cell(row, 4).value
        player2 = ws.cell(row, 5).value
        attend2 = ws.cell(row, 6).value
        if rank and team and player1 and player2:
            if attend1 == -1 or attend2 == -1:
                continue
            teams[team] = {}
            if attend1 == 1 and attend2 == 1:
                teams[team]['ready'] = 1
            else:
                teams[team]['ready'] = 0
            teams[team]['players'] = f'{player1} & {player2}'
            teams[team]['crt5_played'] = False
            teams[team]['score_rows'] = []
            teams[team]['rounds'] = []
            teams[team]['wait_time'] = 1

            if rank == 'wb':
                teams[team]['group'] = 'Women & Boys'
                teams[team]['children'] = True
                groups[4].append(team)
            elif rank == 'g':
                teams[team]['group'] = 'Girls'
                teams[team]['children'] = True
                groups[5].append(team)
            elif isinstance(rank, float) or isinstance(rank, int):
                team_ranks.append([int(rank), team])
                teams[team]['children'] = False
            else:
                print(f'Error in sheet {ATTENDANCE_SHEET} at row {row}, unknown rank "{rank}"')
                exit(1)

    team_ranks.sort(key=lambda k: k[0])
    for [rank, team] in team_ranks:
        rank -= 1
        groups[rank % 4].append(team)
        if (rank % 4) == 0:
            teams[team]['group'] = 'Group A'
        elif (rank % 4) == 1:
            teams[team]['group'] = 'Group B'
        elif (rank % 4) == 2:
            teams[team]['group'] = 'Group C'
        elif (rank % 4) == 3:
            teams[team]['group'] = 'Group D'

    for group in groups:
        matches = combinations(group, 2)
        for match in matches:
            quali_matches.append(list(match))

    random.shuffle(quali_matches)

    return teams, groups, quali_matches

def choose_matches(teams, quali_matches):
    rounds = [None for i in range(COURTS_COUNT)]
    available_rounds = []
    selected_teams = []

    for [left, right] in quali_matches:
        entry = [left, right, teams[left]['ready'], teams[right]['ready'], teams[left]['wait_time'], teams[right]['wait_time']]
        available_rounds.append(entry)

    inx = MENS_COURT_COUNT - 1
    if available_rounds:
        available_rounds.sort(key=lambda k: (k[2]+k[3], k[4]+k[5], k[4], k[5]), reverse=True)
        for i in range(len(available_rounds)):
            left = available_rounds[i][0]
            right = available_rounds[i][1]
            #print(f'avail {left} {right} {available_rounds[i][2]} {available_rounds[i][3]}')
            if left in selected_teams or right in selected_teams:
                continue
            if available_rounds[i][4] == 0 or available_rounds[i][5] == 0:
                continue

            if teams[left]['children'] and teams[right]['children']:
                if not rounds[CHILD_COURT_INX]:
                    rounds[CHILD_COURT_INX] = [left, right]
                    selected_teams.append(left)
                    selected_teams.append(right)
            else:
                if not rounds[CRT5_COURT_INX] and \
                    not teams[left]['crt5_played'] and not teams[right]['crt5_played']:
                    rounds[CRT5_COURT_INX] = [left, right]
                    selected_teams.append(left)
                    selected_teams.append(right)
                elif (inx >= 0):
                    rounds[inx] = [left, right]
                    selected_teams.append(left)
                    selected_teams.append(right)
                    inx -= 1

    return rounds

def update_chosen_teams(teams, quali_matches, rounds):

    selected_teams = []
    for court, match in enumerate(rounds):
        if match:
            teams[match[0]]['wait_time'] = 0
            teams[match[1]]['wait_time'] = 0

            if court == CRT5_COURT_INX:
                teams[match[0]]['crt5_played'] = True
                teams[match[1]]['crt5_played'] = True

            selected_teams.append(match[0])
            selected_teams.append(match[1])

            quali_matches.remove(match)

    for team in teams:
        if team not in selected_teams:
            teams[team]['wait_time'] += 1


def update_schedule(wb, teams, quali_rounds):

    ws = wb[SCHEDULE_SHEET]
    ws.delete_rows(1, ws.max_row)

    schedule_row = [""]
    for i in range(COURTS_COUNT):
        schedule_row.append(f'Court - {i+1}')
        schedule_row.append("")
        schedule_row.append("")
    ws.append(schedule_row)

    for seq, rounds in enumerate(quali_rounds):
        schedule_row = []
        for court, match in enumerate(rounds):
            if court == 0:
                schedule_row.append("")
            if match:
                schedule_row.append(match[0])
                schedule_row.append(match[1])
                schedule_row.append("")
            else:
                schedule_row.append("")
                schedule_row.append("")
                schedule_row.append("")
        ws.append(schedule_row)

    row = ws.max_row + 4
    for col in range(3, len(quali_rounds) + 3):
        ws.cell(row, col, f'Round - {col-2}')

    row = ws.max_row + 1
    for team in teams:
        ws.cell(row, 1, teams[team]['players'])
        ws.cell(row, 2, team)
        rounds = teams[team]['rounds']
        for (seq, court) in rounds:
            ws.cell(row, 2+seq, f'Court - {court}')
        row += 1



def update_score_sheet(wb, groups, teams, quali_rounds):
    ws = wb[SCORE_SHEET]
    ws.delete_rows(1, ws.max_row)
    title = ["players1", "team1", "score", "points", "", "points", "score", "team2", "players2"]
    ws.append(title)
    row = 2

    for seq, rounds in enumerate(quali_rounds):
        ws.append([])
        row += 1
        ws.append([f'Round - {seq+1}'])
        row += 1
        for match in rounds:
            row_entries = ["" for x in title]
            if match:
                left = match[0]
                right = match[1]
                row_entries[0] = teams[left]['players']
                row_entries[1] = left
                left_points_eqn = f'=if({LEFT_SCORE_COL}{row} > {RIGHT_SCORE_COL}{row},{WINNER_POINTS},"")'
                row_entries[3] = left_points_eqn
                right_points_eqn = f'=if({LEFT_SCORE_COL}{row} < {RIGHT_SCORE_COL}{row},{WINNER_POINTS},"")'
                row_entries[5] = right_points_eqn
                row_entries[7] = right
                row_entries[8] = teams[right]['players']
                teams[left]['score_rows'].append(('left', row))
                teams[right]['score_rows'].append(('right', row))
                ws.append(row_entries)
                row += 1
    ws.append([])
    ws.append([])

    for team in teams:
        points_sum_eqn = '=sum('
        score_sum_eqn = '=sum('
        for (side, row) in teams[team]['score_rows']:
            if side == 'left':
                points_sum_eqn += f'{LEFT_POINTS_COL}{row}, '
                score_sum_eqn += f'{LEFT_SCORE_COL}{row}, '
            elif side == 'right':
                points_sum_eqn += f'{RIGHT_POINTS_COL}{row}, '
                score_sum_eqn += f'{RIGHT_SCORE_COL}{row}, '
        points_sum_eqn = points_sum_eqn[:-2] + ')'
        score_sum_eqn = score_sum_eqn[:-2] + ')'
        teams[team]['points_sum_eqn'] = points_sum_eqn
        teams[team]['score_sum_eqn'] = score_sum_eqn

    for group in groups:
        ws.append([teams[group[0]]['group']])
        row_entries = ["" for i in range(7)]
        for team in group:
            row_entries[0] = teams[team]['players']
            row_entries[1] = team
            row_entries[2] = teams[team]['points_sum_eqn']
            row_entries[6] = teams[team]['score_sum_eqn']
            ws.append(row_entries)
        ws.append([])

    return


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f'Argument error\nUsage: {sys.argv[0]} <excel_file>')
        exit(1)

    excel_file = sys.argv[1]
    wb = openpyxl.load_workbook(filename=excel_file)
    teams, groups, quali_matches = retrive_teams(wb)
    quali_rounds = []
    while (True):
        rounds = choose_matches(teams, quali_matches)
        if not any(rounds):
            break
        update_chosen_teams(teams, quali_matches, rounds)
        quali_rounds.append(rounds)

    for seq, rounds in enumerate(quali_rounds):
        for court, match in enumerate(rounds):
            if match:
                teams[match[0]]['rounds'].append((seq+1, court+1))
                teams[match[1]]['rounds'].append((seq+1, court+1))
                print(f'({match[0]:6} {match[1]:6}), ', end='')
            else:
                print(f'({"":6} {"":6}), ', end='')
        print()

    update_schedule(wb, teams, quali_rounds)

    for team in teams:
        rounds = teams[team]['rounds']
        gaps = []
        for i in range(1, len(rounds)):
            gaps.append(rounds[i][0] - rounds[i-1][0])

        print(f'{team}: min {min(gaps)} max {max(gaps)} rounds {" ".join([str(x) for x in rounds])}')
    update_score_sheet(wb, groups, teams, quali_rounds)
    wb.save(excel_file)
