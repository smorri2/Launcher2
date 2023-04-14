#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
import sys

from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime, date


# Third party imports
from typing import Type
import xlsxwriter


# local application imports


# SGM Shared Module imports
# sys.path.insert(1, 'C:/Users/kap3309/OneDrive - Kansas City Life Insurance/PythonDev/Modules')
sys.path.append('C:/Users/kap3309/OneDrive - Kansas City Life Insurance/PythonDev/Modules')
from sgmCsvFileReader import get_csv_file_data
from kclGetFASTProjectInfo import FASTProjectInfo, SprintInfoRec, TeamRec
from kclGetCsvJiraStoryData import CsvJiraStoryData, JiraStoryRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************

@dataclass()
class InputData:
    sprint_info: SprintInfoRec = None
    fast_teams: list[TeamRec] = field(default_factory=list)
    jira_stories: list[JiraStoryRec] = field(default_factory=list)
    success: bool = False


@dataclass
class MetricsRowData:
    label: str
    num_stories: int = 0
    points: int = 0


@dataclass
class StatusTypes:
    done: MetricsRowData
    uat: MetricsRowData
    qa: MetricsRowData
    dev: MetricsRowData
    sdev: MetricsRowData
    tgrm: MetricsRowData
    bgrm: MetricsRowData
    backlog: MetricsRowData


@dataclass
class CategoryTypes:
    new: MetricsRowData
    carryover: MetricsRowData
    unplanned: MetricsRowData


@dataclass
class PriorityTypes:
    highest: MetricsRowData
    high: MetricsRowData
    medium: MetricsRowData
    low: MetricsRowData
    lowest: MetricsRowData


@dataclass
class MetricsData:
    status: StatusTypes
    pi_plan: StatusTypes
    team: list[MetricsRowData]
    category: CategoryTypes
    priority: PriorityTypes
    success: bool = False


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
    def_fmt = None
    table_label_fmt = None


# ******************************************************************************
# ******************************************************************************
# # * Functions
# ******************************************************************************
# ******************************************************************************

# ==============================================================================
def get_input_data():

    fast_project_info: FASTProjectInfo
    input_data = InputData()

    # get the Program Increment to process from the user via console input
    sprint_to_process = get_sprint_to_process()

    print('\n  Begin Getting Input Data ')

    # Get sprint_info for sprint_to_process (eg: name, start_date, end_date, program_increment)
    success = get_project_info_for_sprint_to_process(sprint_to_process, input_data)
    if success:
        # Get the FAST Jira Story data from the 'Jira - Sprint Data.csv' file
        input_data.jira_stories = \
            CsvJiraStoryData(Path.cwd() / 'Input files' / 'SGM - Jira - FAST Sprint Data (Jira).csv').stories
        if input_data.jira_stories is not None:
            input_data.success = True
            print('  Success Getting Input Data')
        else:
            print('\n   *** Error Jira - Sprint Data.csv ***')
            print('  *** Error Getting Input Data ***')
    else:
        print('\n    *** Error getting Sprint information from FAST Project Info.xlsx ***')
        print('  *** Error Getting Input Data ***')

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
def get_project_info_for_sprint_to_process(sprint_to_process: str, input_data: InputData) -> bool:

    success = False
    # Get the FAST Project Info for All Sprints, Team Members and Program Increments
    fast_proj_info = FASTProjectInfo(Path.cwd() / 'Input files' / 'FAST Project Info.xlsx')
    if fast_proj_info is not None:
        # get info about the sprint_to_process from fast_project_info
        input_data.sprint_info = fast_proj_info.get_sprint_info(sprint_to_process)
        if input_data.sprint_info is not None:
            input_data.fast_teams = fast_proj_info.get_fast_teams()
            if len(input_data.fast_teams) > 0:
                success = True

    return success


# ==============================================================================
def get_jira_stories_for_sprint() -> list:

    columns = ['Issue key',
               'Summary',
               'Status',
               'Assignee',
               'Priority',
               'Custom field (Story Points)',
               'Sprint',
               'Labels',
               'Issue Type',
               'Created']
    csv_file_path = Path.cwd() / 'Input files' / 'SGM - Jira - FAST Sprint Data (Jira).csv'
    progress_bar_msg = '      Reading Jira Sprint Data'
    csv_file_data = get_csv_file_data(columns, csv_file_path, progress_bar_msg)

    return csv_file_data


# ==============================================================================
def build_sprint_metrics(input_data: InputData) -> MetricsData:

    print('\n  Begin Building Sprint Metrics')

    # initialize the table data for all tables to be written later to the Metrics tab of the report
    metrics_data = init_metrics_data()

    pi_plan_name = input_data.sprint_info.program_increment
    sprint_data = input_data.sprint_info # get the PI Plan name for the sprint
    for story in input_data.jira_stories:
        update_status_metrics_tbl(metrics_data.status, story.status, story.points)
        update_pi_plan_metrics_tbl(metrics_data.pi_plan, pi_plan_name, story.labels, story.status, story.points)
        update_completed_by_team_metrics_tbl(metrics_data.team, input_data.fast_teams, story.assignee, story.status,
                                             story.points)
        update_story_category_metrics_tbl(metrics_data.category,input_data.sprint_info.name, story.created,
                                          input_data.sprint_info.start_date, story.sprints, story.points)
        update_story_priority_metrics_tbl(metrics_data.priority, story.priority, story.points)
    metrics_data.success = True
    print('\n  Success Building Sprint Metrics')

    return metrics_data


# ==============================================================================
def init_metrics_data() -> MetricsData:

    # initialize Metrics data
    status_metrics = init_status_metrics()
    pi_plan_metrics = init_pi_plan_metrics()
    teams_metrics = []
    category_metrics = init_category_metrics()
    priority_metrics = init_priority_metrics()
    metrics_tbl_data = MetricsData(status_metrics, pi_plan_metrics, teams_metrics, category_metrics, priority_metrics)

    return metrics_tbl_data


# ==============================================================================
def init_status_metrics() -> StatusTypes:
    done = MetricsRowData('Done')
    uat = MetricsRowData('UAT')
    qa = MetricsRowData('QA')
    dev = MetricsRowData('Development')
    sdev = MetricsRowData('Selected for Development')
    tgrm = MetricsRowData('Tech Grooming')
    bgrm = MetricsRowData('Business Grooming')
    backlog = MetricsRowData('Backlog')
    status_metrics = StatusTypes(done, uat, qa, dev, sdev, tgrm, bgrm, backlog)

    return status_metrics


# ==============================================================================
def init_pi_plan_metrics() -> StatusTypes:
    done = MetricsRowData('Done')
    uat = MetricsRowData('UAT')
    qa = MetricsRowData('QA')
    dev = MetricsRowData('Development')
    sdev = MetricsRowData('Selected for Development')
    tgrm = MetricsRowData('Tech Grooming')
    bgrm = MetricsRowData('Business Grooming')
    backlog = MetricsRowData('Backlog')
    pi_plan_metrics = StatusTypes(done, uat, qa, dev, sdev, tgrm, bgrm, backlog)

    return pi_plan_metrics


# ==============================================================================
def init_category_metrics() -> CategoryTypes:
    new = MetricsRowData('New')
    carryover = MetricsRowData('Carry Over')
    unplanned = MetricsRowData('Unplanned')
    category_metrics = CategoryTypes(new, carryover, unplanned)

    return category_metrics


# ==============================================================================
def init_priority_metrics() -> PriorityTypes:
    highest = MetricsRowData('Highest')
    high = MetricsRowData('High')
    medium = MetricsRowData('Medium')
    low = MetricsRowData('Low')
    lowest = MetricsRowData('Lowest')
    priority_metrics = PriorityTypes(highest, high, medium, low, lowest)

    return priority_metrics


# ==============================================================================
def update_status_metrics_tbl(status_tbl: StatusTypes, story_status: str, story_points: int) -> None:
    update_metrics_tbl_with_cur_story_data(status_tbl, story_status, story_points)
    return None


# ==============================================================================
def update_metrics_tbl_with_cur_story_data(metrics_tbl_to_update: StatusTypes, story_status: str, story_points: int) -> None:
    match story_status:
        case 'Done':
            update_metrics_row_data(metrics_tbl_to_update.done, story_points)
        case 'UAT':
            update_metrics_row_data(metrics_tbl_to_update.uat, story_points)
        case 'QA':
            update_metrics_row_data(metrics_tbl_to_update.qa, story_points)
        case 'Development':
            update_metrics_row_data(metrics_tbl_to_update.dev, story_points)
        case 'Selected for Development':
            update_metrics_row_data(metrics_tbl_to_update.sdev, story_points)
        case 'Tech Grooming':
            update_metrics_row_data(metrics_tbl_to_update.tgrm, story_points)
        case 'Business Grooming':
            update_metrics_row_data(metrics_tbl_to_update.bgrm, story_points)
        case 'Backlog':
            update_metrics_row_data(metrics_tbl_to_update.backlog, story_points)

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def update_metrics_row_data(metrics_row_to_update: MetricsRowData, num_points_in) -> None:
    metrics_row_to_update.num_stories += 1
    metrics_row_to_update.points += num_points_in
    return None


# ==============================================================================
def update_pi_plan_metrics_tbl(pi_plan_tbl: StatusTypes, pi_plane_name: str, story_labels: str, story_status: str,
                               story_points: int) -> None:
    if pi_plane_name in story_labels:
        match story_status:
            case 'Done':
                update_metrics_row_data(pi_plan_tbl.done, story_points)
            case 'UAT':
                update_metrics_row_data(pi_plan_tbl.uat, story_points)
            case 'QA':
                update_metrics_row_data(pi_plan_tbl.qa, story_points)
            case 'Development':
                update_metrics_row_data(pi_plan_tbl.dev, story_points)
            case 'Selected for Development':
                update_metrics_row_data(pi_plan_tbl.sdev, story_points)
            case 'Tech Grooming':
                update_metrics_row_data(pi_plan_tbl.tgrm, story_points)
            case 'Business Grooming':
                update_metrics_row_data(pi_plan_tbl.bgrm, story_points)
            case 'Backlog':
                update_metrics_row_data(pi_plan_tbl.backlog, story_points)

    return None


# ==============================================================================
def update_completed_by_team_metrics_tbl(team_tbl: list[MetricsRowData], fast_teams: list[TeamRec], story_assignee: str,
                                         story_status: str, story_points: int) -> None:
    if story_status == 'Done':
        assignee_team = get_assignee_team(fast_teams, story_assignee)
        team_tbl_updated = False
        for cur_team in team_tbl:
            if cur_team.label == assignee_team:
                cur_team.num_stories += 1
                cur_team.points += story_points
                team_tbl_updated = True
                break
        if not team_tbl_updated:
            new_team = MetricsRowData(assignee_team, 1, story_points)
            team_tbl.append(new_team)

    return None


# ==============================================================================
def get_assignee_team(fast_teams: list[TeamRec], story_assignee: str) -> str:
    assignee_team = ''
    for cur_team in fast_teams:
        if story_assignee in cur_team.members:
            assignee_team = cur_team.name
            break

    return assignee_team


# ==============================================================================
def update_story_category_metrics_tbl(category_tbl: CategoryTypes, sprint_name: str, story_created: datetime,
                                      sprint_start_date: datetime, story_sprints: str, story_points: int) -> None:

    if story_created > sprint_start_date:
        update_metrics_row_data(category_tbl.unplanned, story_points)
    else:
        cur_sprint_num = int(sprint_name[12:])
        prev_sprint = sprint_name[:12] + str(cur_sprint_num - 1)
        if prev_sprint in story_sprints:
            update_metrics_row_data(category_tbl.carryover, story_points)
        else:
            update_metrics_row_data(category_tbl.new, story_points)

    return None


# ==============================================================================
def update_story_priority_metrics_tbl(priority_tbl: PriorityTypes, story_priority: str, story_points: int) -> None:
    match story_priority:
        case 'Highest':
            update_metrics_row_data(priority_tbl.highest, story_points)
        case 'High':
            update_metrics_row_data(priority_tbl.high, story_points)
        case 'Medium':
            update_metrics_row_data(priority_tbl.medium, story_points)
        case 'Low':
            update_metrics_row_data(priority_tbl.low, story_points)
        case 'Lowest':
            update_metrics_row_data(priority_tbl.lowest, story_points)

    return None


# ==============================================================================
def create_jira_sprint_report_spreadsheet(metrics_data: MetricsData, input_data: InputData) -> None:
    print('\n  Creating Sprint Metrics Report spreadsheet')

    # create the spreadsheet workbook
    relative_path = 'Output files/' + date.today().strftime("%y-%m-%d") + ' ' + input_data.sprint_info.name \
           + ' Sprint Report.xlsx'
    workbook = xlsxwriter.Workbook(relative_path)
    # create the cell formatting options for the workbook
    cell_formats = create_cell_formatting_options(workbook)

    write_the_metrics_tab_to_spreadsheet(workbook, cell_formats, metrics_data)

    write_the_sprint_stories_tab_to_spreadsheet(workbook, cell_formats, input_data.jira_stories)

    workbook.close()
    print('  Completed PI Planning Metrics Spreadsheet')

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
    cell_fmt.def_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'text_wrap': 1})
    cell_fmt.table_label_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'font_size': 14})

    return cell_fmt


# ==============================================================================
def write_the_metrics_tab_to_spreadsheet(workbook: xlsxwriter.Workbook, cell_fmts: Type[CellFormats],
                                         metrics_data: MetricsData) -> None:
    print('      ** Writing Sprint Metrics spreadsheet tab')

    #  Add the Metrics spreadsheet tab to the workbook
    metrics_ws = workbook.add_worksheet('Sprint Metrics')

    # Set the column widths and default cell formatting for the Metrics worksheet
    create_metrics_ws_column_layout(metrics_ws, cell_fmts)

    write_the_status_metrics_to_ws(metrics_ws, cell_fmts, metrics_data.status)

    write_the_pi_plan_metrics_to_ws(metrics_ws, cell_fmts, metrics_data.pi_plan)

    write_the_team_metrics_to_ws(metrics_ws, cell_fmts, metrics_data.team)

    write_the_category_metrics_to_ws(metrics_ws, cell_fmts, metrics_data.category)

    write_the_priority_metrics_to_ws(metrics_ws, cell_fmts, metrics_data.priority)

    return None


# ==============================================================================
def create_metrics_ws_column_layout(metrics_ws, cell_fmts: Type[CellFormats]) -> None:
    # Set the column widths and default cell formatting for the Metrics tab
    metrics_ws.set_column('A:A', 10)
    metrics_ws.set_column('B:B', 27, cell_fmts.center_fmt)
    metrics_ws.set_column('C:E', 20, cell_fmts.center_fmt)
    metrics_ws.set_column('F:F', 10, cell_fmts.center_fmt)
    metrics_ws.set_column('G:G', 27, cell_fmts.center_fmt)
    metrics_ws.set_column('H:J', 20, cell_fmts.center_fmt)
    metrics_ws.set_column('K:K', 10, cell_fmts.center_fmt)

    return None


# ==============================================================================
def write_the_status_metrics_to_ws(metrics_ws, cell_fmts: Type[CellFormats], status: StatusTypes) -> None:
    print('         Writing Sprint Metrics Table to Metrics Worksheet')

    # Calculate total number of story points in sprint
    num_story_points_total = status.done.points + status.uat.points + status.qa.points + status.dev.points \
        + status.sdev.points + status.tgrm.points + status.bgrm.points + status.backlog.points

    # Calculate percentages
    percent_done = status.done.points / num_story_points_total
    percent_uat = status.uat.points / num_story_points_total
    percent_qa = status.qa.points / num_story_points_total
    percent_dev = status.dev.points / num_story_points_total
    percent_sdev = status.sdev.points / num_story_points_total
    percent_tgrm = status.tgrm.points / num_story_points_total
    percent_bgrm = status.bgrm.points / num_story_points_total
    percent_backlog = status.backlog.points / num_story_points_total

    # Create the table data
    table_data = [
        ['Done', status.done.num_stories, status.done.points, percent_done],
        ['UAT', status.uat.num_stories, status.uat.points, percent_uat],
        ['QA', status.qa.num_stories, status.qa.points, percent_qa],
        ['Development', status.dev.num_stories, status.dev.points, percent_dev],
        ['Selected for Development', status.sdev.num_stories, status.sdev.points, percent_sdev],
        ['Tech Grooming', status.tgrm.num_stories, status.tgrm.points, percent_tgrm],
        ['Business Grooming', status.bgrm.num_stories, status.bgrm.points, percent_bgrm],
        ['Backlog', status.backlog.num_stories, status.backlog.points, percent_backlog]
    ]

    # ******************************************************************
    # Set Status Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    status_tbl = calc_table_starting_and_ending_cells(2, 'B', 'E', 8, True)

    metrics_ws.add_table(status_tbl,
                         {'name': 'Status_Table',
                          'style': 'Table Style Medium 2',
                          'autofilter': False,
                          'first_column': True,
                          'data': table_data,
                          'total_row': 1,
                          'columns': [
                              {'header': 'Sprint Stories by Status', 'total_string': 'Totals',
                               'format': cell_fmts.left_fmt},
                              {'header': '# of Stories', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '# of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '% of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.percent_fmt}]
                          })

    return None


# ==============================================================================
def write_the_pi_plan_metrics_to_ws(metrics_ws, cell_fmts: Type[CellFormats], pi_plane: StatusTypes) -> None:
    print('         Writing PI Plan Metrics Table to Metrics Worksheet')

    # Calculate total number of story points in sprint
    num_story_points_total = pi_plane.done.points + pi_plane.uat.points + pi_plane.qa.points + pi_plane.dev.points \
        + pi_plane.sdev.points + pi_plane.tgrm.points + pi_plane.bgrm.points + pi_plane.backlog.points

    # Calculate percentages
    percent_done = pi_plane.done.points / num_story_points_total
    percent_uat = pi_plane.uat.points / num_story_points_total
    percent_qa = pi_plane.qa.points / num_story_points_total
    percent_dev = pi_plane.dev.points / num_story_points_total
    percent_sdev = pi_plane.sdev.points / num_story_points_total
    percent_tgrm = pi_plane.tgrm.points / num_story_points_total
    percent_bgrm = pi_plane.bgrm.points / num_story_points_total
    percent_backlog = pi_plane.backlog.points / num_story_points_total

    # Create the table data
    table_data = [
        ['Done', pi_plane.done.num_stories, pi_plane.done.points, percent_done],
        ['UAT', pi_plane.uat.num_stories, pi_plane.uat.points, percent_uat],
        ['QA', pi_plane.qa.num_stories, pi_plane.qa.points, percent_qa],
        ['Development', pi_plane.dev.num_stories, pi_plane.dev.points, percent_dev],
        ['Selected for Development', pi_plane.sdev.num_stories, pi_plane.sdev.points, percent_sdev],
        ['Tech Grooming', pi_plane.tgrm.num_stories, pi_plane.tgrm.points, percent_tgrm],
        ['Business Grooming', pi_plane.bgrm.num_stories, pi_plane.bgrm.points, percent_bgrm],
        ['Backlog', pi_plane.backlog.num_stories, pi_plane.backlog.points, percent_backlog]
    ]

    # ******************************************************************
    # Set pi_plane Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    pi_plan_tbl = calc_table_starting_and_ending_cells(14, 'B', 'E', 8, True)

    metrics_ws.add_table(pi_plan_tbl,
                         {'name': 'PI_Plan_Table',
                          'style': 'Table Style Medium 2',
                          'autofilter': False,
                          'first_column': True,
                          'data': table_data,
                          'total_row': 1,
                          'columns': [
                              {'header': 'PI Plan Stories by Status', 'total_string': 'Totals',
                               'format': cell_fmts.left_fmt},
                              {'header': '# of Stories', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '# of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '% of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.percent_fmt}]
                          })

    return None


# ==============================================================================
def write_the_team_metrics_to_ws(metrics_ws, cell_fmts: Type[CellFormats], teams:list[MetricsRowData]) -> None:

    print('         Writing Team Metrics Table to Metrics Worksheet')

    # Calculate total number of story points for all of the Teams combined
    total_team_points = 0
    if teams:
        for cur_team in teams:
            total_team_points += cur_team.points

        # Create the table data
        table_data = []
        for cur_team in teams:
            team_percentage = cur_team.points/ total_team_points
            new_team = [cur_team.label, cur_team.num_stories, cur_team.points, team_percentage]
            table_data.append(new_team)

        # ******************************************************************
        # Set Teams Table starting and ending cells.
        # params are (top_row, left_column, right_column, num_data_rows, Total_row)
        # ******************************************************************
        teams_tbl = calc_table_starting_and_ending_cells(26, 'B', 'E', len(teams), True)

        metrics_ws.add_table(teams_tbl,
                             {'name': 'Teams_Table',
                              'style': 'Table Style Medium 2',
                              'autofilter': False,
                              'first_column': True,
                              'data': table_data,
                              'total_row': 1,
                              'columns': [
                                  {'header': 'Done Stories by Team', 'total_string': 'Totals',
                                   'format': cell_fmts.left_fmt},
                                  {'header': '# of Stories', 'total_function': 'sum',
                                   'format': cell_fmts.right_fmt},
                                  {'header': '# of Story Points', 'total_function': 'sum',
                                   'format': cell_fmts.right_fmt},
                                  {'header': '% of Story Points', 'total_function': 'sum',
                                   'format': cell_fmts.percent_fmt}]
                              })

    return None


# ==============================================================================
def write_the_category_metrics_to_ws(metrics_ws, cell_fmts: Type[CellFormats], categories: CategoryTypes) -> None:

    print('         Writing Category Metrics Table to Metrics Worksheet')

    # Calculate total number of story points for all categories combined
    total_category_points = categories.new.points + categories.carryover.points + categories.unplanned.points

    # Calculate percentages
    percent_new = categories.new.points / total_category_points
    percent_carryover = categories.carryover.points / total_category_points
    percent_unplanned = categories.unplanned.points / total_category_points

    # Create the table data
    table_data = [
        ['New', categories.new.num_stories, categories.new.points, percent_new],
        ['Carryover', categories.carryover.num_stories, categories.carryover.points, percent_carryover],
        ['Unplanned', categories.unplanned.num_stories, categories.unplanned.points, percent_unplanned]
    ]

    # ******************************************************************
    # Set categories Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    categories_tbl = calc_table_starting_and_ending_cells(2, 'G', 'J', 3, True)

    metrics_ws.add_table(categories_tbl,
                         {'name': 'Category_Table',
                          'style': 'Table Style Medium 2',
                          'autofilter': False,
                          'first_column': True,
                          'data': table_data,
                          'total_row': 1,
                          'columns': [
                              {'header': 'Stories by Category', 'total_string': 'Totals',
                               'format': cell_fmts.left_fmt},
                              {'header': '# of Stories', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '# of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '% of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.percent_fmt}]
                          })

    return None


# ==============================================================================
def write_the_priority_metrics_to_ws(metrics_ws, cell_fmts: Type[CellFormats], priorities: PriorityTypes) -> None:

    print('         Writing Category Metrics Table to Metrics Worksheet')

    # Calculate total number of story points for all priorities combined
    total_priority_points = priorities.highest.points + priorities.high.points + priorities.medium.points\
                            + priorities.low.points + priorities.lowest.points

    # Calculate percentages
    percent_highest = priorities.highest.points / total_priority_points
    percent_high = priorities.high.points / total_priority_points
    percent_medium = priorities.medium.points / total_priority_points
    percent_low = priorities.low.points / total_priority_points
    percent_lowest = priorities.lowest.points / total_priority_points

    # Create the table data
    table_data = [
        ['Highest', priorities.highest.num_stories, priorities.highest.points, percent_highest],
        ['High', priorities.high.num_stories, priorities.high.points, percent_high],
        ['Medium', priorities.medium.num_stories, priorities.medium.points, percent_medium],
        ['Low', priorities.low.num_stories, priorities.low.points, percent_medium],
        ['Lowest', priorities.lowest.num_stories, priorities.lowest.points, percent_medium]
    ]

    # ******************************************************************
    # Set priorities Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    priorities_tbl = calc_table_starting_and_ending_cells(9, 'G', 'J', 5, True)

    metrics_ws.add_table(priorities_tbl,
                         {'name': 'Priority_Table',
                          'style': 'Table Style Medium 2',
                          'autofilter': False,
                          'first_column': True,
                          'data': table_data,
                          'total_row': 1,
                          'columns': [
                              {'header': 'Stories by Priority', 'total_string': 'Totals',
                               'format': cell_fmts.left_fmt},
                              {'header': '# of Stories', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '# of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.right_fmt},
                              {'header': '% of Story Points', 'total_function': 'sum',
                               'format': cell_fmts.percent_fmt}]
                          })

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


# ==============================================================================
def write_the_sprint_stories_tab_to_spreadsheet(workbook: xlsxwriter.Workbook, cell_fmts: Type[CellFormats],
                                                jira_stories: list[JiraStoryRec]) -> None:
    print('      ** Writing Jira Stories Detail spreadsheet tab')

    #  Add the Metrics spreadsheet tab to the workbook
    detail_ws = workbook.add_worksheet('Jira Stories Detail')

    # Set the column widths and default cell formatting for the Metrics worksheet
    create_sprint_stories_column_layout(detail_ws, cell_fmts)

    write_the_sprint_stories_to_ws(detail_ws, cell_fmts, jira_stories)

    return None


# ==============================================================================
def create_sprint_stories_column_layout(jira_stories_ws, cell_fmts: Type[CellFormats]) -> None:

    # Setup Jira table layout
    jira_stories_ws.set_column('A:B', 15, cell_fmts.center_fmt)
    jira_stories_ws.set_column('C:C', 80, cell_fmts.center_fmt)
    jira_stories_ws.set_column('D:F', 25, cell_fmts.center_fmt)
    jira_stories_ws.set_column('G:G', 12, cell_fmts.center_fmt)
    jira_stories_ws.set_column('H:H', 16, cell_fmts.center_fmt)
    jira_stories_ws.set_column('I:I', 22, cell_fmts.center_fmt)
    jira_stories_ws.set_column('J:J', 40, cell_fmts.center_fmt)
    jira_stories_ws.set_column('K:K', 80, cell_fmts.center_fmt)

    return None


# ==============================================================================
def write_the_sprint_stories_to_ws(jira_data_ws, cell_fmts: Type[CellFormats], jira_stories: list[JiraStoryRec]) -> None:

    # ******************************************************************
    # Set Jira Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    jira_data_tbl = calc_table_starting_and_ending_cells(1, 'A', 'K', len(jira_stories), False)
    table_data = []
    for cur_story in jira_stories:
        sprints = ';'.join([str(sprint) for sprint in cur_story.sprints])
        new_row_data = [cur_story.key,
                        cur_story.type,
                        cur_story.summary,
                        cur_story.status,
                        cur_story.assignee,
                        cur_story.test_assignee,
                        cur_story.priority,
                        cur_story.points,
                        cur_story.created.strftime("%m/%d/%Y, %H:%M:%S"),
                        ';'.join([str(sprint) for sprint in cur_story.sprints]),  # list comprehension to convert list
                        cur_story.labels]
        table_data.append(new_row_data)

    jira_data_ws.add_table(jira_data_tbl,
                           {'name': 'Jira_Data_table',
                            'style': 'Table Style Medium 2',
                            'autofilter': True,
                            'first_column': False,
                            'data': table_data,
                            'columns': [{'header': 'Issue Key', 'format': cell_fmts.center_fmt},
                                        {'header': 'Issue Type', 'format': cell_fmts.center_fmt},
                                        {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                        {'header': 'Status', 'format': cell_fmts.left_fmt},
                                        {'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                        {'header': 'Test Assignee', 'format': cell_fmts.left_fmt},
                                        {'header': 'Priority', 'format': cell_fmts.left_fmt},
                                        {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                        {'header': 'Created', 'format': cell_fmts.center_fmt},
                                        {'header': 'Sprint', 'format': cell_fmts.left_fmt},
                                        {'header': 'Labels', 'format': cell_fmts.left_fmt}]
                            })

    return None


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************
def main():

    print('\nBegin Create FAST Sprint Report')
    input_data = get_input_data()
    if input_data.success:
        sprint_metrics = build_sprint_metrics(input_data)
        if sprint_metrics.success:
            create_jira_sprint_report_spreadsheet(sprint_metrics, input_data)

    print('\nEnd Create FAST Sprint Report')

    return None


if __name__ == "__main__":
    main()
