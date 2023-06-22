#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from pathlib import Path
# from dataclasses import dataclass
from typing import Type


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


@dataclass()
class InputData:
    sprint_info: SprintRec = None
    team_info: list[TeamRec] = None
    jira_stories: FastStoryData = None
    success: bool = False


@dataclass
class AssigneeStoryRec:
    data: FastStoryRec
    cur_assignee: str


class AssigneeDataRec:
    def __init__(self, assignee_in, jira_story_in):
        self.assignee_name: str = assignee_in
        self.stories: list[AssigneeStoryRec] = jira_story_in
        self.ws = None


@dataclass
class AssigneeTeamsData:
    kcl_assignees: list[AssigneeDataRec] = field(default_factory=list)
    it_assignees: list[AssigneeDataRec] = field(default_factory=list)
    actuarial_assignees: list[AssigneeDataRec] = field(default_factory=list)
    verisk_assignees: list[AssigneeDataRec] = field(default_factory=list)


@dataclass
class CellFormats:
    metrics_ws_fmt = None
    left_fmt = None
    left_bold_fmt = None
    header_left_fmt = None
    header_center_fmt = None
    right_fmt = None
    percent_fmt = None
    percent_center_fmt = None
    center_fmt = None
    center_red_fmt = None
    blue_fmt = None
    def_fmt = None
    table_label_fmt = None


class AssigneeSS:
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


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************
def create_standup_assignees_spreadsheet():

    print('\n\nStart Create Sprint Standup Assignee Spreadsheet')

    input_data = get_input_data()
    if input_data.success:
        assignee_data = process_jira_stories(input_data.jira_stories, input_data.team_info)
        create_fast_standup_assignees_ss(assignee_data)
    print('\nEnd Create Sprint Standup Assignee Spreadsheet')

    return None


# ==============================================================================
def get_input_data():

    input_data = InputData()

    # get the Sprint to process from the user via console input
    sprint_to_process = get_sprint_to_process()

    print('\n  Begin Getting Input Data ')

    # Get Fast Team info, Teams and Members from the FastTeamInfo.csv spreadsheet
    fast_teams_info = FASTTeams(Path.cwd())
    if fast_teams_info is not None:
        input_data.team_info = fast_teams_info.teams
        # Get FAST Sprint info, start date and end date, from the FastSprintInfo.csv spreadsheet
        input_data.sprint_info = FASTSprints(Path.cwd())
        if input_data.sprint_info is not None:
            input_data.sprint_info = input_data.sprint_info.get_sprint_info(sprint_to_process)
            # Get the FAST Jira Story data for the sprint being processed
            jql_query = create_jql_query(input_data.sprint_info.name[5:])
            input_data.jira_stories = FastStoryData(jql_query).stories
            if input_data.jira_stories is not None:
                input_data.success = True
                print('   Success Getting Input Data')
            else:
                print('   *** Error getting Sprint Info from SGM - Jira - FAST Sprint Data (Jira).csv')
        else:
            print('   *** Error getting Sprint Info from FastSprintInfo.csv')
    else:
        print('   *** Error getting Team Info from FastTeamInfo.csv')

    return input_data


# ==============================================================================
def get_sprint_to_process() -> str:
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
        sprint_to_process = '2023 FASTR1i69'

    return sprint_to_process


# ==============================================================================
def create_jql_query(sprint_name) -> str:
    project = 'project = "FAST" AND '
    sprint = 'Sprint = ' + sprint_name + ' AND '
    story_type = 'Type in (Bug, Story, Task, Sub-task) AND '
    status = 'Status in (UAT, QA, Development, "Selected for Development", "Tech Grooming", "Business Grooming", ' \
             'Backlog) '
    order_by = 'ORDER BY Key'
    jql_query = project + sprint + story_type + status + order_by

    return jql_query


# ==============================================================================
# ==============================================================================
# * Functions
# ==============================================================================
# ==============================================================================


# ==============================================================================
def process_jira_stories(jira_stories: FastStoryData, teams_info: list[TeamRec]) -> AssigneeTeamsData:

    print('   Begin Processing Assignee Stories')
    assignee_teams = AssigneeTeamsData()
    for cur_jira_story in jira_stories:
        current_assignee = get_current_assignee(cur_jira_story)
        assignee_story_rec = AssigneeStoryRec(cur_jira_story, current_assignee)
        cur_assignee_team = get_cur_assignee_team(current_assignee, teams_info)

        match cur_assignee_team:
            case 'Verisk':
                update_assignee_data(assignee_teams.verisk_assignees, current_assignee, assignee_story_rec)
            case 'App Systems':
                update_assignee_data(assignee_teams.it_assignees, 'IT', assignee_story_rec)
            case 'Data Services':
                update_assignee_data(assignee_teams.it_assignees, 'IT', assignee_story_rec)
            case 'FAST Support':
                update_assignee_data(assignee_teams.it_assignees, 'IT', assignee_story_rec)
            case 'Actuarial':
                update_assignee_data(assignee_teams.actuarial_assignees, 'Actuarial', assignee_story_rec)
            case _:
                update_assignee_data(assignee_teams.kcl_assignees, current_assignee, assignee_story_rec)

    return assignee_teams


# ==============================================================================
def get_current_assignee(jira_story_rec: FastStoryRec) -> str:
    if jira_story_rec.test_assignee == '':
        cur_assignee = jira_story_rec.assignee
    else:
        match jira_story_rec.status:
            case 'UAT':
                cur_assignee = jira_story_rec.test_assignee
            case 'QA':
                cur_assignee = jira_story_rec.test_assignee
            case _:
                cur_assignee = jira_story_rec.assignee

    return cur_assignee


# ==============================================================================
def get_cur_assignee_team(assignee: str, teams_info: list[TeamRec]) -> str:

    assignee_team = ''
    for cur_team in teams_info:
        if assignee in cur_team.members:
            assignee_team = cur_team.name
            break
    return assignee_team


# ==============================================================================
def update_assignee_data(assignee_data: list[AssigneeDataRec], assignee_to_update: str,
                         assignee_story: AssigneeStoryRec) -> None:

    assignee_found = False
    if len(assignee_data) > 0:
        for cur_assignee_rec in assignee_data:
            if cur_assignee_rec.assignee_name == assignee_to_update:
                assignee_found = True
                cur_assignee_rec.stories.append(assignee_story)
                break
    if not assignee_found:
        new_assignee = AssigneeDataRec(assignee_to_update, [assignee_story])
        assignee_data.append(new_assignee)

    return None


# ==============================================================================
def create_fast_standup_assignees_ss(assignee_data: AssigneeTeamsData) -> None:

    print('\n   Creating FAST Standup Assignee spreadsheet')
    assignee_ss = AssigneeSS()
    cur_date = datetime.now().strftime("%y-%m-%d")
    # create the spreadsheet workbook and formats for the spreadsheet
    assignee_ss.workbook = xlsxwriter.Workbook('Output files/' + cur_date + ' Sprint Standup Assignees.xlsx')
    cell_formats = create_cell_formatting_options(assignee_ss.workbook)
    # Sort the assignees so that they are displayed in Alphabetic order
    assignee_data.verisk_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee_name)
    assignee_data.it_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee_name)
    assignee_data.actuarial_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee_name)
    assignee_data.kcl_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee_name)
    # Write the Verisk team members stories to the worksheet
    print('      Writing Verisk Assignees to Spreadsheet')
    for cur_assignee in assignee_data.verisk_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.data.status, reverse=True)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee_name)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)
    # Write an empty worksheet named Verisk Questions that will be used as a prompt to have KCLife team ask any
    # questions they have before we let Verisk team leave Standup meeting
    assignee_ws = assignee_ss.workbook.add_worksheet('Verisk Questions')
    print('      Writing IT Assignees to Spreadsheet')
    for cur_assignee in assignee_data.it_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.cur_assignee)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee_name)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)
    # Write the Actuarial team members stories to the worksheet
    print('      Writing Actuarial Assignees to Spreadsheet')
    for cur_assignee in assignee_data.actuarial_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.cur_assignee)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee_name)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)
    # Write the remaining kcl team members stories to the worksheet
    for cur_assignee in assignee_data.kcl_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.data.status)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee_name)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)

    assignee_ss.workbook.close()
    print('   Done creating FAST Standup Assignee spreadsheet')

    return None


# ==============================================================================
def create_cell_formatting_options(workbook) -> Type[CellFormats]:
    # create predefined cell_formats to be used for cells in the workbook
    cell_fmt = CellFormats
    cell_fmt.metrics_ws_fmt = workbook.add_format({'font_name': 'Calibri', 'align': 'center', 'font_size': 12})
    cell_fmt.left_fmt = workbook.add_format({'align': 'left', 'indent': 1})
    cell_fmt.left_bold_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'indent': 1})
    cell_fmt.header_left_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'indent': 1, 'font_size': 12})
    cell_fmt.header_center_fmt = workbook.add_format({'align': 'center', 'bold': 1, 'font_size': 12})
    cell_fmt.right_fmt = workbook.add_format({'align': 'right', 'indent': 8})
    cell_fmt.percent_fmt = workbook.add_format({'align': 'right', 'indent': 8, 'num_format': '0%'})
    cell_fmt.percent_center_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
    cell_fmt.center_fmt = workbook.add_format({'align': 'center'})
    cell_fmt.center_red_fmt = workbook.add_format({'align': 'center', 'font_color': 'red', 'bold': 1})
    cell_fmt.blue_fmt = workbook.add_format({'align': 'center', 'font_color': 'blue', 'bold': 1})
    cell_fmt.def_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'text_wrap': 1})
    cell_fmt.table_label_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'font_size': 14})

    return cell_fmt


# ==============================================================================
def create_assignee_ws_column_layout(worksheet, cell_fmts: Type[CellFormats]) -> None:
    # Set the column widths and default cell formatting for the Metrics tab
    # Setup Jira table layout
    worksheet.set_column('A:A', 20, cell_fmts.center_fmt)
    worksheet.set_column('B:B', 16, cell_fmts.center_fmt)
    worksheet.set_column('C:C', 12, cell_fmts.center_fmt)
    worksheet.set_column('D:D', 80, cell_fmts.center_fmt)
    worksheet.set_column('E:F', 20, cell_fmts.center_fmt)
    worksheet.set_column('G:H', 12, cell_fmts.center_fmt)

    return None


# ==============================================================================
def write_assignee_stories_to_ws(assignee_ws, cell_fmts: Type[CellFormats],
                                 assignee_stories: list[AssigneeStoryRec]) -> None:

    # ******************************************************************
    # Set Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    assignee_story_tbl = calc_table_starting_and_ending_cells(1, 'A', 'H', len(assignee_stories), False)
    table_data = []
    for cur_story in assignee_stories:
        new_row_data = [cur_story.cur_assignee,
                        cur_story.data.status,
                        cur_story.data.issue_key,
                        cur_story.data.summary,
                        cur_story.data.assignee,
                        cur_story.data.test_assignee,
                        cur_story.data.priority,
                        cur_story.data.points]
        table_data.append(new_row_data)

    table_name = assignee_ws.name.replace(' ', '_')
    assignee_ws.add_table(assignee_story_tbl,
                          {'name': table_name,
                           'style': 'Table Style Medium 2',
                           'autofilter': True,
                           'first_column': False,
                           'data': table_data,
                           'columns': [{'header': 'Current Assignee', 'format': cell_fmts.left_fmt},
                                       {'header': 'Status', 'format': cell_fmts.left_fmt},
                                       {'header': 'Issue Key', 'format': cell_fmts.center_fmt},
                                       {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                       {'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                       {'header': 'Test Assignee', 'format': cell_fmts.left_fmt},
                                       {'header': 'Priority', 'format': cell_fmts.left_fmt},
                                       {'header': 'Story Points', 'format': cell_fmts.center_fmt}]
                           })

    assignee_ws.conditional_format('A2:A40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="QA"',
                                    'format': cell_fmts.blue_fmt})
    assignee_ws.conditional_format('F2:F40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="QA"',
                                    'format': cell_fmts.blue_fmt})
    assignee_ws.conditional_format('A2:A40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="UAT"',
                                    'format': cell_fmts.blue_fmt})
    assignee_ws.conditional_format('F2:F40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="UAT"',
                                    'format': cell_fmts.blue_fmt})
    assignee_ws.conditional_format('B2:B40',
                                   {'type': 'text',
                                    'criteria': 'containsText',
                                    'value': 'QA',
                                    'format': cell_fmts.blue_fmt})
    assignee_ws.conditional_format('B2:B40',
                                   {'type': 'text',
                                    'criteria': 'containsText',
                                    'value': 'UAT',
                                    'format': cell_fmts.blue_fmt})
    return None


# ==============================================================================
def calc_table_starting_and_ending_cells(top_row: int, left_col: str, right_col: str, num_data_rows: int,
                                         total_row: bool) -> str:
    """
    Calculates the starting and ending cells of an Excel table

    param top_row: int top row number of the table
    param left_col: str left most column of the table
    param right_col: str - right most column the table
    param num_data_rows: int - number of rows of data in the table
    param total_row: bool - does the table have a Totals Row in addition to the data rows
    return: str containing table coordinates to pass to the xlsxwriter add_table function
    """
    top_left_cell = left_col + str(top_row)
    if total_row:
        bot_right_cell = right_col + str(top_row + num_data_rows + 1)
    else:
        bot_right_cell = right_col + str(top_row + num_data_rows)
    table_coordinates = top_left_cell + ':' + bot_right_cell

    return table_coordinates


if __name__ == "__main__":
    create_standup_assignees_spreadsheet()
