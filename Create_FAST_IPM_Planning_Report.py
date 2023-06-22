#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from dataclasses import dataclass, field
from pathlib import Path


# Third party imports
import xlsxwriter


# local file imports


# SGM Shared Module imports
from kclFastSharedDataClasses import *
from kclGetFastTeams import FASTTeams
from kclGetFastSprints import FASTSprints
from kclGetFastStoryDataJiraAPI import FastStoryData, FastStoryRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************

@dataclass
class InputData:
    sprint_info: SprintRec = None
    team_info: FASTTeams = None
    prev_sprint: str = ''
    fast_teams: list[TeamRec] = field(default_factory=list)
    jira_stories: FastStoryData = None
    success: bool = False


@dataclass
class IpmPlanningRec(FastStoryRec):
    carry_over_story: str = 'N'


class AssigneesRec:
    def __init__(self, assignee_in, jira_story_in, story_points_in):
        self.assignee: str = assignee_in
        self.stories: list[IpmPlanningRec] = jira_story_in
        self.total_points: int = story_points_in
        self.ws = None


class IpmPlanningSS:
    def __init__(self):
        self.workbook = None
        self.totals_ws = None
        self.header_fmt = None
        self.left_fmt = None
        self.left_bold_fmt = None
        self.left_lv2_fmt = None
        self.right_fmt = None
        self.center_fmt = None
        self.center_red_fmt = None
        self.center_orange_fmt = None
        self.center_green_fmt = None
        self.percent_fmt = None
        self.header_fmt = None
        self.last_row_fmt = None
        self.totals_fmt = None
        self.assignee_data: list = []


# ==============================================================================
# ==============================================================================
# * Functions
# ==============================================================================
# ==============================================================================


# ==============================================================================
def get_input_data():

    input_data = InputData()

    # get the Program Increment to process from the user via console input
    sprint_to_process = get_program_increment_to_process()

    print('\n  Begin Getting Input Data')

    # Get Fast Team info, Teams and Members from the FastTeamInfo.csv spreadsheet
    input_data.team_info = FASTTeams(Path.cwd())
    if input_data.team_info is not None:
        input_data.sprint_info = get_sprint_info(sprint_to_process)
        if input_data.sprint_info is not None:
            input_data.prev_sprint = get_prev_sprint_name(sprint_to_process)
            # Get the FAST Jira Story data for the sprint being processed
            jql_query = create_jql_query(input_data.sprint_info.name[5:])
            input_data.jira_stories = FastStoryData(JIRA_USER, JIRA_TOKEN, JIRA_OPTIONS, jql_query).stories
            if input_data.jira_stories is not None:
                input_data.success = True
                print('  Success Getting Input Data')
            else:
                print('   *** Error getting Sprint Info from SGM - Jira - FAST Sprint Data (Jira).csv')
        else:
            print('   *** Error getting Sprint Info from FastSprintInfo.csv')
    else:
        print('   *** Error getting Team Info from FastTeamInfo.csv')

    return input_data


# ==============================================================================
def get_sprint_info(sprint_to_process: str) -> SprintRec:
    # Get FAST Sprint info, start date and end date, from the FastSprintInfo.csv spreadsheet
    fast_sprints_info = FASTSprints(Path.cwd())
    if fast_sprints_info is not None:
        sprint_info = fast_sprints_info.get_sprint_info(sprint_to_process)
    else:
        sprint_info = None
    return sprint_info


# ==============================================================================
def get_prev_sprint_name(sprint_to_process: str) -> str:
    sprint_num = int(sprint_to_process[12:])
    prev_sprint_name = sprint_to_process[:12] + str(sprint_num - 1)
    return prev_sprint_name


# ==============================================================================
def create_jql_query(sprint_name) -> str:
    project = 'project = "FAST" AND '
    sprint = 'Sprint = ' + sprint_name + ' AND '
    story_type = 'Type in (Bug, Story, Task, Sub-task) AND '
    status = 'Status in (UAT, QA, Development, "Selected for Development", "Tech Grooming", ' \
             '"Business Grooming", Backlog) '
    order_by = 'ORDER BY Key'
    jql_query = project + sprint + story_type + status + order_by

    return jql_query


# ==============================================================================
def get_program_increment_to_process() -> str:
    sprint_to_process: str = ''
    valid_input: bool = False

    debug: bool = False
    if not debug:
        while not valid_input:
            print('\n')
            print('   ************************************************')
            print('   **                                           ***')
            print('   **    Enter Sprint Number to Report On       ***')
            print('   **                                           ***')
            print('   ************************************************')

            user_input = input('\n   Enter Sprint Number to process (two digits only) ==> ')
            if user_input.isdecimal():  # Verify that the user input was a number
                sprint_num = int(user_input)
                if 39 < sprint_num < 100:
                    sprint_to_process = '2023 FASTR1i' + str(sprint_num)
                    valid_input = True
                else:
                    print('\n\n   Invalid Sprint Number, valid Sprint Numbers are between 40 & 99 inclusive')
            else:
                print('\n\n   Invalid option Selected, enter two digit sprint number only')
    else:  # When debugging hard code the sprint name to avoid having to get input from console
        sprint_to_process = '2023 FASTR1i70'

    return sprint_to_process


# ==============================================================================
def process_ipm_planning_data(input_data: InputData) -> list:
    stories_by_assignee = []
    for cur_story_rec in input_data.jira_stories:
        new_planning_rec = copy_cur_story_rec_data(cur_story_rec)
        # check if this is a carry over story and set carry_over_story field to 'Y' if True
        if input_data.prev_sprint in cur_story_rec.sprints:
            new_planning_rec.carry_over_story = 'Y'
        update_planning_spreadsheet_assignees_stories(stories_by_assignee, new_planning_rec)

    return stories_by_assignee


# ==============================================================================
def copy_cur_story_rec_data(cur_story_rec: FastStoryRec) -> IpmPlanningRec:

    new_planning_rec = IpmPlanningRec()
    new_planning_rec.issue_type = cur_story_rec.issue_type
    new_planning_rec.issue_key = cur_story_rec.issue_key
    new_planning_rec.summary = cur_story_rec.summary
    new_planning_rec.status = cur_story_rec.status
    new_planning_rec.created = cur_story_rec.created
    new_planning_rec.priority = cur_story_rec.priority
    new_planning_rec.assignee = cur_story_rec.assignee
    new_planning_rec.sprints = cur_story_rec.sprints
    new_planning_rec.labels = cur_story_rec.labels
    new_planning_rec.points = cur_story_rec.points

    return new_planning_rec


# ===============================================================================
def update_planning_spreadsheet_assignees_stories(assignees_list: list[AssigneesRec],
                                                  planning_rec: IpmPlanningRec) -> None:
    new_assignee = True
    if assignees_list:  # Check to see if the assignees list of assignees is empty
        if planning_rec.assignee == '':
            planning_rec.assignee = 'Unassigned'
        # iterate through the assignees list checking to see if this assignee already has a record in the list
        for cur_assignees_rec in assignees_list:
            if cur_assignees_rec.assignee == planning_rec.assignee:
                # this assignee already exists, so update the data for this assignee
                new_assignee = False  # assignee found, so not a new assignee for the list
                cur_assignees_rec.stories.append(planning_rec)  # add story to this assignees list of stories
                cur_assignees_rec.total_points += planning_rec.points
                break
    if new_assignee:
        # The assignees list is empty or this is a new assignee for the assignees list, so create a new assignees rec
        story_list = [planning_rec]
        new_assignees_rec = AssigneesRec(planning_rec.assignee, story_list, planning_rec.points)
        assignees_list.append(new_assignees_rec)

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def create_sprint_report_spreadsheet(stories_by_assignee: list[AssigneesRec], input_data: InputData) -> IpmPlanningSS:

    print('\n   Creating IPM Planning spreadsheet')
    # create the spreadsheet workbook and formats for the IPM Planning spreadsheet
    ipm_planning_ss = create_ss_workbook_and_formats(input_data.sprint_info.name)

    # change order of assignee's in list to move high priorty assignee's to front of list
    ipm_planning_ss.assignee_data = update_order_of_assignees(stories_by_assignee, input_data.team_info)

    # Setup the All Assignees worksheet tab to hold the totals by Assignee
    ipm_planning_ss.totals_ws = ipm_planning_ss.workbook.add_worksheet('All Assignees')
    ipm_planning_ss.totals_ws.set_column('A:A', 20)
    ipm_planning_ss.totals_ws.write('A1', 'Assignee', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.set_column('B:C', 14)
    ipm_planning_ss.totals_ws.write('B1', 'Initial Story Points', ipm_planning_ss.header_fmt)
    ipm_planning_ss.totals_ws.write('C1', 'Final Story Points', ipm_planning_ss.header_fmt)

    # create the worksheet and table layouts for the Metrics tab in the IPM Planning spreadsheet
    for cur_assignee in ipm_planning_ss.assignee_data:
        cur_assignee.ws = ipm_planning_ss.workbook.add_worksheet(cur_assignee.assignee)

        # Write out the worksheet header
        cur_assignee.ws.set_column('A:A', 12)
        cur_assignee.ws.write('A1', 'Key', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('B:B', 9, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('B1', 'Issue Type', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('C:C', 70, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('C1', 'Summary', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('D:D', 18, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('D1', 'Assignee', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('E:E', 25, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('E1', 'Status', ipm_planning_ss.header_fmt)
        cur_assignee.ws.set_column('F:F', 12, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('F1', 'Priority', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('G:J', 12, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('G1', 'Initial Story Points', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('H1', 'Carryover Story', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('I1', 'Remaining Story Points', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('J1', 'Final Story Points', ipm_planning_ss.header_fmt)

    return ipm_planning_ss


# ==============================================================================
def create_ss_workbook_and_formats(sprint_to_plan: str) -> IpmPlanningSS:

    # create the IPM Planning spreadsheet data structure and then create spreadsheet workbook
    ipm_planning_ss = IpmPlanningSS()
    ipm_planning_ss.workbook = xlsxwriter.Workbook('Output files/' + sprint_to_plan + ' IPM Planning.xlsx')

    # add predefined formats to be used for formatting cells in the spreadsheet
    ipm_planning_ss.left_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'indent': 1
    })
    ipm_planning_ss.left_bold_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    ipm_planning_ss.left_lv2_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'indent': 4
    })
    ipm_planning_ss.right_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'right',
        'indent': 6
    })
    ipm_planning_ss.center_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
    })
    ipm_planning_ss.center_red_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF0000'
    })
    ipm_planning_ss.center_orange_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF8000'
    })
    ipm_planning_ss.center_green_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
        'bold': 1,
        'font_color': '#00CC66'
    })
    ipm_planning_ss.percent_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'right',
        'indent': 6,
        'num_format': '0%'
    })
    ipm_planning_ss.header_fmt = ipm_planning_ss.workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 13,
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
        'font_size': 12,
        'align': 'center',
        'bottom': 6
    })
    ipm_planning_ss.totals_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 13,
        'align': 'center',
        'bold': 1,
    })

    return ipm_planning_ss


# ==============================================================================
def create_assignee_worksheet(ipm_planning_ss: IpmPlanningSS) -> None:

    print('      ** Writing IPM Planning spreadsheet tab')
    ipm_planning_ss.data_ws = ipm_planning_ss.workbook.add_worksheet('IMP Planning')

    # Setup Details table layout
    ipm_planning_ss.data_ws.set_column('A:B', 14, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('C:C', 80, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('D:E', 15, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('F:F', 20, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('G:G', 18, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('H:H', 20, ipm_planning_ss.center_fmt)

    # ******************************************************************
    # Set Sprint Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    ipm_planning_ss.detail_table = calc_table_starting_and_ending_cells(1, 'A', 'H', len(jira_sprint_data) - 1)

    return None


# ==============================================================================
def calc_table_starting_and_ending_cells(top_row: int, left_col, right_col, num_data_rows) -> str:
    top_left_cell = left_col + str(top_row)
    bot_right_cell = right_col + str(top_row + num_data_rows + 1)
    table_coordinates = top_left_cell + ':' + bot_right_cell

    return table_coordinates


# ==============================================================================
def update_order_of_assignees(stories_by_assignee: list[AssigneesRec], fast_teams_info: FASTTeams) -> list[AssigneesRec]:

    updated_assignee_list = []
    # Find the high priority assignee's in the stories_by_assignee list so they can be moved to front
    # of the list in the next step.  This will allow them to be the first tabs in the IPM Spreadsheet
    high_priority_list = []
    normal_priority_list = []
    for cur_assignee in stories_by_assignee:
        assignee_team = fast_teams_info.get_team_of_member(cur_assignee.assignee)
        match assignee_team:
            case 'Verisk':
                high_priority_list.append(cur_assignee)
            case '':
                high_priority_list.append(cur_assignee)
            case _:
                normal_priority_list.append(cur_assignee)

    high_priority_list.sort(key=lambda assignee_rec: assignee_rec.assignee, reverse=True)
    normal_priority_list.sort(key=lambda assignee_rec: assignee_rec.assignee)

    updated_assignee_list = high_priority_list + normal_priority_list

    return updated_assignee_list


# ==============================================================================

def write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> None:
    bottom_row = len(ipm_planning_ss.assignee_data)
    ws_row = 0
    for cur_assignees_rec in ipm_planning_ss.assignee_data:
        ws_row += 1
        ipm_planning_ss.totals_ws.write(ws_row, 0, cur_assignees_rec.assignee, ipm_planning_ss.left_fmt)
        total_row = len(cur_assignees_rec.stories) + 3
        initial_points_total_loc = "='" + cur_assignees_rec.assignee + "'!G" + str(total_row)
        final_points_total_loc = "='" + cur_assignees_rec.assignee + "'!J" + str(total_row)
        if ws_row == bottom_row:
            cell_fmt = ipm_planning_ss.last_row_fmt
        else:
            cell_fmt = ipm_planning_ss.center_fmt
        ipm_planning_ss.totals_ws.write(ws_row, 1, initial_points_total_loc, cell_fmt)
        ipm_planning_ss.totals_ws.write(ws_row, 2, final_points_total_loc, cell_fmt)

    total_row = '=sum(G2:G' + str(ws_row + 1) + ')'
    ipm_planning_ss.totals_ws.write(ws_row + 1, 1, '=sum(B2:B' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)
    ipm_planning_ss.totals_ws.write(ws_row + 1, 2, '=sum(C2:C' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)

    return None


# ==============================================================================

def write_ipm_planning_data_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> None:

    for cur_assignees_rec in ipm_planning_ss.assignee_data:
        cur_assignees_rec.stories.sort(key=lambda planning_rec: planning_rec.carry_over_story, reverse=True)
        bottom_row = len(cur_assignees_rec.stories)
        ws_row = 0  # leave a empty row above the first row of data for easier manual insertion during IPM
        for cur_story in cur_assignees_rec.stories:
            ws_row += 1
            print(cur_story.issue_key, cur_story.assignee)
            cur_assignees_rec.ws.write(ws_row, 0, cur_story.issue_key, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 1, cur_story.issue_type, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 2, cur_story.summary, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 3, cur_story.assignee, ipm_planning_ss.left_fmt)
            # cur_assignees_rec.ws.write(ws_row, 4, cur_story.status, ipm_planning_ss.center_fmt)
            write_status_to_spreadsheet(cur_assignees_rec.ws, ws_row, cur_story.status, ipm_planning_ss)
            cur_assignees_rec.ws.write(ws_row, 5, cur_story.priority, ipm_planning_ss.center_fmt)
            write_story_points_to_spreadsheet(cur_assignees_rec.ws, ws_row, cur_story.points, ipm_planning_ss)

            # Write formula's to cell to sum Initial Story points, Remainging story points, Final Story points
            cell_fmt = ipm_planning_ss.center_fmt
            cur_assignees_rec.ws.write(ws_row, 7, cur_story.carry_over_story, cell_fmt)
            # formula to calculate remaining story points, if col H = 'Y' then it's a carryover story so return the
            # initial story points found in col G, if not then it is a new story so return an empty string
            remaining_points_fml = '=IF(H' + str(ws_row + 1) + '="Y", G' + str(ws_row + 1) + ', "")'
            cur_assignees_rec.ws.write(ws_row, 8, remaining_points_fml, cell_fmt)
            # formula to calculate fina story points, if col I is an empty string "" then it's a new story return the
            # initial story points, else it's a carryover story so return the remaining story points in col I
            final_points_fml = '=IF(I' + str(ws_row + 1) + '="", G' + str(ws_row + 1) + ', I' + str(ws_row + 1) + ')'
            cur_assignees_rec.ws.write(ws_row, 9, final_points_fml, cell_fmt)

        # leave an empty row between last story and totals row for easier story insertion during IPM
        ws_row += 1
        test = (' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ')
        cur_assignees_rec.ws.write_row(ws_row, 0, test, ipm_planning_ss.last_row_fmt)

        cur_assignees_rec.ws.write(ws_row + 1, 6, '=sum(G2:G' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)
        cur_assignees_rec.ws.write(ws_row + 1, 9, '=sum(J2:J' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)

    return None


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
    if story_points == 0:
        cell_fmt = ipm_planning_ss.center_red_fmt
    else:
        cell_fmt = ipm_planning_ss.center_fmt
    ws.write(ws_row, 6, story_points, cell_fmt)

    return None


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************
def create_fast_ipm_planning_spreadsheet():

    print('\n\nStart Create IPM Planning Spreadsheet')

    ipm_planning_ss = IpmPlanningSS()

    input_data = get_input_data()
    if input_data is not None:
        ipm_planning_ss.assignee_data = process_ipm_planning_data(input_data)

    ipm_planning_ss = create_sprint_report_spreadsheet(ipm_planning_ss.assignee_data, input_data)

    write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss)
    write_ipm_planning_data_to_spreadsheet(ipm_planning_ss)
    ipm_planning_ss.workbook.close()

    print('\nEnd Create IPM Planning Spreadsheet')

    return None


if __name__ == "__main__":
    create_fast_ipm_planning_spreadsheet()
