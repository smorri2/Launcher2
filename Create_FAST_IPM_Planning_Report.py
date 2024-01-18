#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from dataclasses import dataclass
from pathlib import Path


# Third party imports
import xlsxwriter


# local file imports


# SGM Shared Module imports
from kclFastSharedDataClasses import *
from kclGetFastInfo import FASTInfoDB, SprintRec
from kclGetFastStoryDataJiraAPI import FastStoryData, FastStoryRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************


@dataclass
class IpmPlanningRec:
    story_data: FastStoryRec
    carry_over_story: str = ''


@dataclass
class AssigneesListRec:
    assignee: str = ''
    stories: list[IpmPlanningRec] = field(default_factory=list)
    points: int = 0
    total_row: int = 0


@dataclass
class TeamListRec:
    team_name: str
    members: list[AssigneesListRec]
    points: int
    total_row: int
    ws = None


@dataclass
class IpmPlanningSS:
    workbook = None
    totals_ws = None
    left_fmt = None
    left_bold_fmt = None
    left_lv2_fmt = None
    right_fmt = None
    center_fmt = None
    center_red_fmt = None
    center_orange_fmt = None
    center_green_fmt = None
    percent_fmt = None
    header_fmt = None
    last_row_fmt = None
    last_row_name_fmt = None
    totals_fmt = None
    assignees: list[AssigneesListRec] = field(default_factory=list)
    teams: list[TeamListRec] = field(default_factory=list)


@dataclass
class InputData:
    fast_info_db: FASTInfoDB = None
    sprint_info: SprintRec = None
    prev_sprint: str = ''
    jira_stories: FastStoryData = None
    success: bool = False


@dataclass
class ProcessedData:
    assignees_list: list[AssigneesListRec]
    teams_list: list[TeamListRec]


# ==============================================================================
# ==============================================================================
# === Functions
# ==============================================================================
# ==============================================================================


# ==============================================================================
# ==============================================================================
# === Main
# ==============================================================================
# ==============================================================================
def create_fast_ipm_planning_spreadsheet():
    print('\n\nStart Create IPM Planning Spreadsheet')

    input_data = get_input_data()
    if input_data.success:
        processed_data = process_ipm_planning_data(input_data)
        if processed_data:
            create_ipm_planning_spreadsheet(processed_data, input_data.sprint_info.name)
    else:
        print('   Error getting Input Data, application terminated')

    print('\nEnd Create IPM Planning Spreadsheet')

    return None


# ==============================================================================
def get_input_data():
    input_data = InputData()

    print('\n  Begin Getting Input Data')

    # Get Fast Team info, Teams and Members from the FastInfo.db sqlite database
    input_data.fast_info_db = FASTInfoDB(Path.cwd())
    if input_data.fast_info_db is not None:
        input_data.sprint_info = input_data.fast_info_db.request_sprint_to_report_on_return_sprint_info()

        # input_data.sprint_info = input_data.fast_info_db.get_sprint_info(sprint_name)
        if input_data.sprint_info is not None:
            # Using the sprint info for the sprint being planned, get the name of the Previous
            # sprint so that we can use it to determine if the current story being processed
            # is a carryover story from the previous sprint.
            input_data.prev_sprint = input_data.fast_info_db.get_prev_sprint_name(input_data.sprint_info.name)

            # Get the FAST Jira Story data for the sprint being processed
            jql_query = create_jql_query(input_data.sprint_info.name[5:])
            input_data.jira_stories = FastStoryData(jql_query).stories
            if input_data.jira_stories is not None:
                input_data.success = True
                print('  Success Getting Input Data')
            else:
                print('   *** Error getting FAST Story Data using Jira API')
        else:
            print('   *** Error getting Sprint Info from FastInfo.db')
    else:
        print('   *** Error accessing FastInfo.db')

    return input_data


# ==============================================================================
def create_jql_query(sprint_name) -> str:
    project = 'project = "FAST" AND '
    sprint = 'Sprint = ' + sprint_name + ' AND '
    story_type = 'Type in (Bug, Story, Task) AND '
    status = 'Status in (UAT, QA, Development, "Selected for Development", "Tech Grooming", ' \
             '"Business Grooming", Backlog) '
    order_by = 'ORDER BY Key'
    jql_query = project + sprint + story_type + status + order_by

    return jql_query


# ==============================================================================
def process_ipm_planning_data(input_data: InputData) -> ProcessedData | None:
    processed_data = None
    assignees = process_jira_stories_for_sprint_to_plan(input_data)
    if assignees:
        teams = process_assignee_list_for_sprint_to_plan(assignees, input_data.fast_info_db)
        if teams:
            processed_data = ProcessedData(assignees, update_order_of_teams(teams))

    return processed_data


# ==============================================================================
def process_jira_stories_for_sprint_to_plan(input_data) -> list[AssigneesListRec]:
    # Create the list of assignee stories that will be used to create the IPM Planning
    # spreadsheet.  Each assignee in the list will have a list of stories assigned
    # to them for the report.
    assignee_list = []

    # loop thru the jira stories in the input_data and update the assignees_list
    # with the new_planning_rec data
    for cur_story_rec in input_data.jira_stories:
        carryover_story = get_carryover_status(input_data.prev_sprint, cur_story_rec.sprints, cur_story_rec.status)
        new_planning_rec = IpmPlanningRec(cur_story_rec, carryover_story)
        # loop thru the assignees in the assignee list and add the current jira
        # story rec to the assignee in the assignee list that matches the assignee
        # in the current jira story.  If assignee for the current jira story is not
        # found in the assignee_list the else clause will create a new assignee list rec
        # and append it to the existing assignee list
        index = 0
        while index < len(assignee_list):
            if cur_story_rec.assignee == assignee_list[index].assignee:
                assignee_list[index].stories.append(new_planning_rec)
                assignee_list[index].points += cur_story_rec.points
                break
            index += 1
        else:
            new_assignee_rec = AssigneesListRec(cur_story_rec.assignee, [new_planning_rec], cur_story_rec.points)
            assignee_list.append(new_assignee_rec)

    return assignee_list


# ==============================================================================
def process_assignee_list_for_sprint_to_plan(assignee_list: list[AssigneesListRec],
                                             fast_info_db: FASTInfoDB) -> list[TeamListRec]:
    # Create the list of teams that will be used to create the IPM Planning
    # spreadsheet.  Each team in the list will have a list of assignees assigned
    # to them for the report.
    teams_list = []

    # loop thru the assignees in the assignees_list and update the teams_list
    # with the new_planning_rec data
    for cur_assignee_rec in assignee_list:
        # get the team name of the cur_assignee_rec.assignee
        assignee_team = fast_info_db.get_assignee_team(cur_assignee_rec.assignee)
        # loop thru the teams in the teams_list and add the current cur_assignee_rec to
        # the team in the teams_list that the assignee in the cur_assignee_rec is a
        # member of.  if team for the current assignee_list assignee is not found in
        # the teams_list the else clause will create a new TeamListRec and append it
        # to the existing teams_list
        index = 0
        while index < len(teams_list):
            if assignee_team == teams_list[index].team_name:
                teams_list[index].members.append(cur_assignee_rec)
                teams_list[index].points += cur_assignee_rec.points
                break
            index += 1
        else:
            new_teams_rec = TeamListRec(assignee_team, [cur_assignee_rec], cur_assignee_rec.points, 0)
            teams_list.append(new_teams_rec)

    return teams_list


# ==============================================================================
def get_carryover_status(prev_sprint: str, sprints_in_story: str, story_status: str) -> str:
    # Check to see if this is a carry over story by seeing if last sprint is
    # in the Sprints list for this story.  If so then set carryover flag to Y
    if prev_sprint in sprints_in_story:
        carry_over_story = 'Y'
    else:
        carry_over_story = 'N'

    return carry_over_story


# ===============================================================================
def create_ipm_planning_spreadsheet(processed_data: ProcessedData,
                                    sprint_name: str) -> None:
    if processed_data is not None:
        ipm_planning_ss = create_spreadsheet(processed_data, sprint_name)
        write_teams_ipm_planning_data_to_spreadsheet(ipm_planning_ss)
        next_row = write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss)
        write_ipm_planning_team_totals_to_spreadsheet(ipm_planning_ss, next_row)
        ipm_planning_ss.workbook.close()

    return None


# ===============================================================================
def create_spreadsheet(processed_data: ProcessedData, sprint_name: str) -> IpmPlanningSS:
    print('\n   Creating IPM Planning spreadsheet')

    # create the spreadsheet workbook and formats for the IPM Planning spreadsheet
    ipm_planning_ss = create_ss_workbook_and_formats(processed_data, sprint_name)

    # Change order of assignee's in the list to move high priority assignee's to front of list
    # Order of list should be:
    # Unassigned - First so that we can assign these to someone and move to their tab
    # Unknown - next, rare but assignee is unknown if they are not in the FastInfo.db
    # Verisk - next so that we can review the Verisk stories and then let them go from IPM
    # iPace - next
    # All the rest

    # Set up the All Assignees worksheet tab to hold the totals by Assignee
    ipm_planning_ss.totals_ws = ipm_planning_ss.workbook.add_worksheet('All Assignees')
    ipm_planning_ss.totals_ws.set_column('A:A', 25)
    ipm_planning_ss.totals_ws.set_column('B:C', 15)

    # create the worksheet and table layouts for each Teams tab in the IPM Planning spreadsheet
    for cur_team in ipm_planning_ss.teams:
        cur_team.ws = ipm_planning_ss.workbook.add_worksheet(cur_team.team_name)

        # Write out the worksheet header
        cur_team.ws.set_column('A:A', 12)  # Issue Key
        cur_team.ws.set_column('B:B', 8, ipm_planning_ss.center_fmt)  # Story Type
        cur_team.ws.set_column('C:C', 80, ipm_planning_ss.center_fmt)  # Summary
        cur_team.ws.set_column('D:D', 19, ipm_planning_ss.center_fmt)  # Assignee
        cur_team.ws.set_column('E:E', 20, ipm_planning_ss.center_fmt)  # Status
        cur_team.ws.set_column('F:F', 11, ipm_planning_ss.center_fmt)  # Priority
        cur_team.ws.set_column('G:J', 11, ipm_planning_ss.center_fmt)  # Initial, Carryover, Remaining, and Final Points

    return ipm_planning_ss


# ==============================================================================
def create_ss_workbook_and_formats(processed_data: ProcessedData, sprint_to_plan: str) -> IpmPlanningSS:
    # create the IPM Planning spreadsheet data structure and then create spreadsheet workbook
    ipm_planning_ss = IpmPlanningSS()
    ipm_planning_ss.assignees = processed_data.assignees_list
    ipm_planning_ss.teams = processed_data.teams_list

    ipm_planning_ss.workbook = xlsxwriter.Workbook('Output files/' + sprint_to_plan + ' IPM Planning.xlsx')

    font_size = 12
    # add predefined formats to be used for formatting cells in the spreadsheet
    ipm_planning_ss.left_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 1
    })
    ipm_planning_ss.left_bold_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    ipm_planning_ss.left_lv2_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 4
    })
    ipm_planning_ss.right_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'right',
        'indent': 6
    })
    ipm_planning_ss.center_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
    })
    ipm_planning_ss.center_red_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF0000'
    })
    ipm_planning_ss.center_orange_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF8000'
    })
    ipm_planning_ss.center_green_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#00CC66'
    })
    ipm_planning_ss.percent_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'right',
        'indent': 6,
        'num_format': '0%'
    })
    ipm_planning_ss.header_fmt = ipm_planning_ss.workbook.add_format({
        'font_name': 'Calibri',
        'font_size': font_size,
        'font_color': 'white',
        'text_wrap': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bold': 1,
        'bg_color': '#4472C4',
        'pattern': 1,
        'border': 1
    })
    ipm_planning_ss.last_row_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bottom': 6
    })
    ipm_planning_ss.last_row_name_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 1,
        'bottom': 6
    })
    ipm_planning_ss.totals_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
    })

    return ipm_planning_ss


# ==============================================================================
def calc_table_starting_and_ending_cells(top_row: int, left_col, right_col, num_data_rows) -> str:
    top_left_cell = left_col + str(top_row)
    bot_right_cell = right_col + str(top_row + num_data_rows + 1)
    table_coordinates = top_left_cell + ':' + bot_right_cell

    return table_coordinates


# ==============================================================================
def update_order_of_teams(teams_list: list[TeamListRec]) -> list[TeamListRec]:
    reordered_teams_list = []

    # Find the Unassigned team in the teams list, if it exists append it to the
    # reordered list and remove it from the teams list
    unassigned = find_team_rec_in_teams_list('Unassigned', teams_list)
    if unassigned is not None:
        reordered_teams_list.append(unassigned)
        teams_list.remove(unassigned)

    # Find the Unknown team in the teams list, if it exists append it to
    # reordered list and remove it from the teams list
    unknown = find_team_rec_in_teams_list('Unknown', teams_list)
    if unknown is not None:
        reordered_teams_list.append(unknown)
        teams_list.remove(unknown)

    # Find the Verisk team in the teams list, if it exists append it to
    # reordered list
    verisk = find_team_rec_in_teams_list('Verisk', teams_list)
    if verisk is not None:
        reordered_teams_list.append(verisk)
        teams_list.remove(verisk)

    # Sort the remaining teams in the teams list and then append to the
    # reordered list
    teams_list.sort(key=lambda teams_list_rec: teams_list_rec.team_name, reverse=True)
    for next_team in teams_list:
        reordered_teams_list.append(next_team)

    return reordered_teams_list


# ==============================================================================
def find_team_rec_in_teams_list(team_to_find, teams_list: list[TeamListRec]) -> TeamListRec | None:
    team_rec = None
    for cur_team_rec in teams_list:
        if cur_team_rec.team_name == team_to_find:
            team_rec = cur_team_rec
            break

    return team_rec


# ==============================================================================
def write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> int:
    bottom_row = len(ipm_planning_ss.assignees)
    ws_row = 0

    # Write the Assignee Totals Header
    ipm_planning_ss.totals_ws.write('A1', 'Assignee', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.write('B1', 'Story Points', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.write('C1', 'IPM Points', ipm_planning_ss.header_fmt)

    # Write the Assignee Totals for each Assignee
    for cur_team in ipm_planning_ss.teams:
        for cur_team_member in cur_team.members:
            ws_row += 1
            initial_points_total_loc = f"='{cur_team.team_name}'!G{str(cur_team_member.total_row)}"
            final_points_total_loc = f"='{cur_team.team_name}'!J{str(cur_team_member.total_row)}"

            if ws_row < bottom_row:
                ipm_planning_ss.totals_ws.write(ws_row, 0, cur_team_member.assignee, ipm_planning_ss.left_fmt)
                ipm_planning_ss.totals_ws.write(ws_row, 1, initial_points_total_loc, ipm_planning_ss.center_fmt)
                ipm_planning_ss.totals_ws.write(ws_row, 2, final_points_total_loc, ipm_planning_ss.center_fmt)
            else:
                ipm_planning_ss.totals_ws.write(ws_row, 0, cur_team_member.assignee, ipm_planning_ss.last_row_name_fmt)
                ipm_planning_ss.totals_ws.write(ws_row, 1, initial_points_total_loc, ipm_planning_ss.last_row_fmt)
                ipm_planning_ss.totals_ws.write(ws_row, 2, final_points_total_loc, ipm_planning_ss.last_row_fmt)

    # Write the formula's to sum the Assignee Totals
    ipm_planning_ss.totals_ws.write(ws_row + 1, 1, '=sum(B2:B' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)
    ipm_planning_ss.totals_ws.write(ws_row + 1, 2, '=sum(C2:C' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)

    return ws_row


# ==============================================================================
def write_ipm_planning_team_totals_to_spreadsheet(ipm_planning_ss: IpmPlanningSS, ws_row: int) -> None:
    ws_row += 5
    header_row = ws_row
    bottom_row = header_row + len(ipm_planning_ss.teams) - 1

    # Write the Assignee Totals Header
    ipm_planning_ss.totals_ws.write(f'A{header_row}', 'Team', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.write(f'B{header_row}', 'Story Points', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.write(f'C{header_row}', 'IPM Points', ipm_planning_ss.header_fmt)

    # Write the Team Totals for each Team
    for cur_team in ipm_planning_ss.teams:
        totals = build_total_initial_points_formula_for_team(cur_team.team_name, cur_team.members)
        if ws_row < bottom_row:
            ipm_planning_ss.totals_ws.write(ws_row, 0, cur_team.team_name, ipm_planning_ss.left_fmt)
            ipm_planning_ss.totals_ws.write(ws_row, 1, totals[0], ipm_planning_ss.center_fmt)
            ipm_planning_ss.totals_ws.write(ws_row, 2, totals[1], ipm_planning_ss.center_fmt)
        else:
            ipm_planning_ss.totals_ws.write(ws_row, 0, cur_team.team_name, ipm_planning_ss.last_row_name_fmt)
            ipm_planning_ss.totals_ws.write(ws_row, 1, totals[0], ipm_planning_ss.last_row_fmt)
            ipm_planning_ss.totals_ws.write(ws_row, 2, totals[1], ipm_planning_ss.last_row_fmt)
        ws_row += 1

    # Write the formula's to sum the Team Totals
    ipm_planning_ss.totals_ws.write(ws_row, 1, f'=sum(B{header_row+1}:B{ws_row})', ipm_planning_ss.totals_fmt)
    ipm_planning_ss.totals_ws.write(ws_row, 2, f'=sum(C{header_row+1}:C{ws_row})', ipm_planning_ss.totals_fmt)

    return None


# ==============================================================================
def build_total_initial_points_formula_for_team(team_name: str, team_members: list[AssigneesListRec]) -> tuple:
    points_formula = f'='
    ipm_points_formula = f'='
    initial_points_col = 'G'
    ipm_points_col = 'J'

    for index, cur_assignee in enumerate(team_members):
        if index < len(team_members) - 1:
            points_formula += f"'{team_name}'!{initial_points_col}{cur_assignee.total_row}+"
            ipm_points_formula += f"'{team_name}'!{ipm_points_col}{cur_assignee.total_row}+"
        else:
            points_formula += f"'{team_name}'!{initial_points_col}{cur_assignee.total_row}"
            ipm_points_formula += f"'{team_name}'!{ipm_points_col}{cur_assignee.total_row}"

    totals = (points_formula, ipm_points_formula)
    return totals


# ==============================================================================
def write_teams_ipm_planning_data_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> None:
    for cur_team in ipm_planning_ss.teams:
        ws_row = 0

        # Sort the members of the team in alphabetic order.  Note: based on first name
        cur_team.members.sort(key=lambda members_rec: members_rec.assignee)

        for cur_team_member in cur_team.members:
            ws_row = write_team_member_header(ws_row, cur_team, ipm_planning_ss)
            ws_row = write_team_member_stories_to_ws(ws_row,
                                                     cur_team,
                                                     cur_team_member,
                                                     ipm_planning_ss)
            ws_row += 1

    return None


# ==============================================================================
def write_team_member_header(cur_row: int, cur_team: TeamListRec, ipm_planning_ss: IpmPlanningSS) -> int:

    cur_team.ws.write(cur_row, 0, 'Key', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 1, 'Type', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 2, 'Summary', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 3, 'Assignee', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 4, 'Status', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 5, 'Priority', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 6, 'Points', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 7, 'Carryover', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 8, 'Remaining', ipm_planning_ss.header_fmt)
    cur_team.ws.write(cur_row, 9, 'IPM Points', ipm_planning_ss.header_fmt)
    cur_row += 1

    return cur_row


# ==============================================================================
def write_team_member_stories_to_ws(cur_row: int,
                                    cur_team: TeamListRec,
                                    assignee: AssigneesListRec,
                                    ipm_ss: IpmPlanningSS) -> int:

    assignee.stories.sort(key=lambda ipm_planning_rec: ipm_planning_rec.carry_over_story, reverse=True)

    print(f'   Writing {assignee.assignee} to {cur_team.team_name}')
    first_story_row = cur_row
    for cur_story in assignee.stories:
        cur_team.ws.write(cur_row, 0, cur_story.story_data.issue_key, ipm_ss.left_fmt)
        cur_team.ws.write(cur_row, 1, cur_story.story_data.issue_type, ipm_ss.left_fmt)
        cur_team.ws.write(cur_row, 2, cur_story.story_data.summary, ipm_ss.left_fmt)
        cur_team.ws.write(cur_row, 3, cur_story.story_data.assignee, ipm_ss.left_fmt)
        write_status_to_spreadsheet(cur_team.ws, cur_row, cur_story.story_data.status, ipm_ss)
        cur_team.ws.write(cur_row, 5, cur_story.story_data.priority, ipm_ss.center_fmt)
        write_story_points_to_spreadsheet(cur_team.ws, cur_row, cur_story.story_data.points, ipm_ss)
        cur_team.ws.write(cur_row, 7, cur_story.carry_over_story, ipm_ss.center_fmt)

        # formula to calculate remaining story points, if col H = 'Y' then it's a carryover story so return the
        # initial story points found in col G, if not then it is a new story so return an empty string
        remaining_points_fml = '=IF(H' + str(cur_row + 1) + '="Y", G' + str(cur_row + 1) + ', "")'
        cur_team.ws.write(cur_row, 8, remaining_points_fml, ipm_ss.center_fmt)
        # formula to calculate fina story points, if col I is an empty string "" then it's a new story return the
        # initial story points, else it's a carryover story so return the remaining story points in col I
        final_points_fml = '=IF(I' + str(cur_row + 1) + '="", G' + str(cur_row + 1) + ', I' + str(cur_row + 1) + ')'
        cur_team.ws.write(cur_row, 9, final_points_fml, ipm_ss.center_fmt)
        cur_row += 1

    # leave an empty row between last story and totals row for easier story insertion during IPM
    blank_row = (' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ')
    cur_team.ws.write_row(cur_row, 0, blank_row, ipm_ss.last_row_fmt)
    cur_row += 1

    # Write the formula's to calculate the total points and IPM Planning points for this team member
    last_story_row = cur_row
    total_points_fml = f'=sum(G{first_story_row}:G{last_story_row})'
    imp_points_fml = f'=sum(J{first_story_row}:J{last_story_row})'
    # Write the totals for team member
    cur_team.ws.write(cur_row, 6, total_points_fml, ipm_ss.totals_fmt)
    cur_team.ws.write(cur_row, 9, imp_points_fml, ipm_ss.totals_fmt)

    # Save the row number of the total row for this team member as it will be used
    # on the assignees tab to display this assignee's total points and IPM Planning points
    assignee.total_row = last_story_row + 1

    # Now increment the row for the next team members story data
    cur_row += 1

    return cur_row


# ==============================================================================
def write_status_to_spreadsheet(ws, ws_row: int, status: str, ipm_planning_ss: IpmPlanningSS) -> None:
    # Highlight status in Red if they are in either Business Grooming or Tech Grooming
    match status:
        case 'Business Grooming':
            cell_fmt = ipm_planning_ss.center_red_fmt
        case 'Tech Grooming':
            cell_fmt = ipm_planning_ss.center_orange_fmt
        case 'Done':
            cell_fmt = ipm_planning_ss.center_green_fmt
        case _:
            cell_fmt = ipm_planning_ss.center_fmt
    ws.write(ws_row, 4, status, cell_fmt)

    return None


# ==============================================================================
def write_story_points_to_spreadsheet(ws, ws_row: int, story_points, ipm_planning_ss: IpmPlanningSS) -> None:
    # Highlight story points in Red if they are currently zero
    if story_points not in [1, 2, 3, 5, 8, 13, 21]:
        cell_fmt = ipm_planning_ss.center_red_fmt
    else:
        cell_fmt = ipm_planning_ss.center_fmt
    ws.write(ws_row, 6, story_points, cell_fmt)

    return None


if __name__ == "__main__":
    create_fast_ipm_planning_spreadsheet()
