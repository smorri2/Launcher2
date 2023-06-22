#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from time import sleep


# Third party imports
from typing import Type
import xlsxwriter
from workdays import networkdays


# local application imports


# SGM Shared Module imports
from kclFastSharedDataClasses import *
from kclGetFastSprints import FASTSprints
from kclGetPIPlannedStoriesData_1 import PIPlannedStoryData, PlannedStoryRec
from kclGetFastStoryDataJiraAPI import FastStoryData, FastStoryRec

# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************


@dataclass
class InputData:
    success: bool = False
    errors: list[str] = field(default_factory=list)
    pi_info: ProgramIncrementRec = None
    pi_planned_stories_data: PIPlannedStoryData = None
    jira_stories: FastStoryData = None


@dataclass
class SprintTotalsRec:
    stories: int = 0
    points: int = 0
    done_stories: int = 0
    done_points: int = 0


@dataclass
class SprintRec:
    sprint_name: str
    planned_totals: SprintTotalsRec
    unplanned_totals: SprintTotalsRec
    combined_totals: SprintTotalsRec


@dataclass
class PlannedStoryDetailRec:
    key: str = ''
    summary: str = ''
    plan_points: int = 0
    cur_points: int = 0
    plan_sprint: str = ''
    cur_sprint: str = ''
    status: str = ''
    team: str = ''


@dataclass
class MetricsData:
    success: bool = False
    sprints: list[SprintRec] = field(default_factory=list)
    planned: list[PlannedStoryDetailRec] = field(default_factory=list)
    removed: list[PlannedStoryRec] = field(default_factory=list)
    unassigned: list[FastStoryRec] = field(default_factory=list)
    unplanned: list[FastStoryRec] = field(default_factory=list)
    jira_detail: list[FastStoryRec] = field(default_factory=list)


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
def get_input_file_data():

    input_data = InputData()
    # get the Program Increment to process from the user via console input
    pi_to_process = get_pi_to_process()
    if pi_to_process != '':

        print('\n  Begin Getting Input Data')
        # Get Program Increment data from the FAST Project Info.xlsx file
        input_data.pi_info = get_program_increment_info(pi_to_process)
        if input_data.pi_info is None:
            input_data.errors.append('\n*** Error Reading FastSprintInfo.csv ***')

        # Get planned stories from PI_Plan_Qx_2023.xlsx file
        input_data.pi_planned_stories_data = get_planned_stories_for_program_increment(pi_to_process)
        if input_data.pi_planned_stories_data is None:
            input_data.errors.append('\n*** Error Reading PI_Plan_Qx_2023 - Planned.xlsx ***')

        # Get the FAST Jira Story data for the Program Increment being processed
        jql_query = create_jql_query(input_data.pi_info.pi_name)
        input_data.jira_stories = FastStoryData(jql_query)
        if input_data.jira_stories is not None:
            input_data.success = True
            print('  Success Getting Input Data')
        else:
            print('   *** Error getting Sprint Info from SGM - Jira - FAST Sprint Data (Jira).csv')

    return input_data


# ==============================================================================
def get_pi_to_process() -> str:
    pi_to_process: str = ''
    valid_input: bool = False

    while not valid_input:
        print('\n')
        print('   ************************************************')
        print('   **                                           ***')
        print('   **    Select the PI to Process               ***')
        print('   **                                           ***')
        print('   **         1 - PI_Plan_Q1_2023               ***')
        print('   **         2 - PI_Plan_Q2_2023               ***')
        print('   **         3 - PI_Plan_Q3_2023               ***')
        print('   **         4 - PI_Plan_Q4_2023               ***')
        print('   **                                           ***')
        print('   ************************************************')

        user_input = input('\n   Enter Number to process  ==> ')
        match user_input:
            case '1':
                pi_to_process = 'PI_Plan_Q1_2023'
                valid_input = True
            case '2':
                pi_to_process = 'PI_Plan_Q2_2023'
                valid_input = True
            case '3':
                pi_to_process = 'PI_Plan_Q3_2023'
                valid_input = True
            case '4':
                pi_to_process = 'PI_Plan_Q4_2023'
                valid_input = True
            case _:
                valid_input = False
                print('\n\n\n   Invalid PI Number, valid PI Numbers are between 1 & 4 inclusive')

    return pi_to_process


# ==============================================================================
def get_program_increment_info(pi_to_process: str) -> ProgramIncrementRec:
    """
    This function:
        Takes a path to the FastSprintInfo.csv file and reads
        the sprint and program increment data into memory so that we can then search the Sprint info
        for the pi_to_process and get the list of sprints in the pi along
        with the start date and end date of the pi.

    :return: A ProgramIncrementRec object containing the data info for all sprints in
        this pi
    """

    print('    Get Program Increment Info from FastSprintInfo.csv file')
    pi_info = None
    fast_sprint_pi_info = FASTSprints(Path.cwd())
    if fast_sprint_pi_info.success:
        pi_info = fast_sprint_pi_info.get_pi_info(pi_to_process)

    return pi_info


# ==============================================================================

def get_planned_stories_for_program_increment(pi_to_process: str) -> PIPlannedStoryData:
    """
    This function:
        creates a PIPlannedStoryData object passing in the path to the
        PI Planned stories spreadsheet for the current Program Increment.
        The PIPlannedStoryData object will read in all the planned story
        data from the PI Planned stories - Q1 spreadsheet.  Later the
        planned stories will be processed to build the PI Planning Status
        report spreadsheet.
    return: A PIPlannedStoryData object containing the data and methods for
        the stories planned for the current Program Increment.
    """
    input_folder = 'Input files'
    filename = pi_to_process + ' - Planned.xlsx'
    planned_story_data = PIPlannedStoryData(Path.cwd() / input_folder / filename)
    sleep(0.05)
    return planned_story_data


# ==============================================================================
def create_jql_query(pi_name: str) -> str:
    project = 'project = "FAST" AND '
    labels = 'labels = ' + pi_name
    order_by = ' ORDER BY key'
    jql_query = project + labels + order_by

    return jql_query


# ==============================================================================
def process_program_increment_data(input_data: InputData) -> MetricsData:
    print('\n  Begin processing input data')
    # Create metrics_data dataclass to hold metrics data from processing input
    metrics_data = MetricsData()

    success = init_metrics_data(metrics_data, input_data.pi_info.pi_sprints)
    if success:
        process_pi_planned_stories(input_data, metrics_data)
        process_unplanned_stories_added_to_pi(input_data, metrics_data)
        metrics_data.success = True
        print('  Completed processing input data')

    return metrics_data


# ==============================================================================
def init_metrics_data(metrics_data: MetricsData, sprints: list) -> bool:
    # setup return flag
    success = False

    # initialize list of sprints in metrics data
    for sprint_name_to_add in sprints:
        new_planned = SprintTotalsRec()
        new_unplanned = SprintTotalsRec()
        new_combined = SprintTotalsRec()
        new_sprint = SprintRec(sprint_name_to_add, new_planned, new_unplanned, new_combined)
        metrics_data.sprints.append(new_sprint)
    if len(metrics_data.sprints) == len(sprints):
        success = True
        # Now that we have verified that we have the same number of sprints as
        # are contained in the PI in the metrics_data.sprints.  Add the additional
        # 'Not Assigned to a Sprint' sprint to keep track of stories that were planned
        # with a sprint but have had the sprint removed.
        not_assigned_sprint = SprintRec('Not Assigned to Sprint', SprintTotalsRec(), SprintTotalsRec(),
                                        SprintTotalsRec())
        metrics_data.sprints.append(not_assigned_sprint)
        print('    Metrics Data initialized')
    else:
        print('    *** Error initializing Metrics Data')

    return success


# ==============================================================================
def process_pi_planned_stories(input_data: InputData, metrics_data: MetricsData) -> None:

    print('    Processing PI Planned Stories')
    for cur_planned_story in input_data.pi_planned_stories_data:
        cur_jira_story = input_data.jira_stories.get_story(cur_planned_story.key)
        if cur_jira_story is not None:
            update_sprint_totals_for_planned_story(cur_planned_story, cur_jira_story, metrics_data)
            update_pi_planned_detail_list(metrics_data.planned, cur_planned_story, cur_jira_story)
        else:
            # If cur_jira_story is None then the Planned story has been removed from the PI,
            # which we want to report on later
            print('     Planned story ' + cur_planned_story.key + ' has been removed from the PI')
            metrics_data.removed.append(cur_planned_story)

    return None


# ==============================================================================
def update_sprint_totals_for_planned_story(planned_story: PlannedStoryRec, jira_story: FastStoryRec,
                                           metric_data: MetricsData) -> None:

    # Check to see if the jira_story.sprints has been unassigned
    if jira_story.sprints != '':
        sprint_to_find = planned_story.sprint
    else:
        sprint_to_find = 'Not Assigned to Sprint'
        metric_data.unassigned.append(jira_story)

    # Loop through the metrics_data list of sprints to find the correct sprint to
    # add this planned jira story points to the totals for that sprint in the PI
    found = False
    for cur_sprint in metric_data.sprints:
        if cur_sprint.sprint_name == sprint_to_find:
            found = True
            # Update the planned totals for the cur_sprint
            cur_sprint.planned_totals.stories += 1
            cur_sprint.planned_totals.points += jira_story.points
            # Update the planned + unplanned totals for the cur_sprint
            cur_sprint.combined_totals.stories += 1
            cur_sprint.combined_totals.points += jira_story.points
            # Check status of Jira story to see if the status is Done
            if jira_story.status == 'Done':
                # Update the Done planned totals for the cur_sprint
                cur_sprint.planned_totals.done_stories += 1
                cur_sprint.planned_totals.done_points += jira_story.points
                # Update the planned + unplanned totals for the cur_sprint
                cur_sprint.combined_totals.done_stories += 1
                cur_sprint.combined_totals.done_points += jira_story.points
            break  # updated the totals for the correct sprint, stop the loop
    if not found:
        # The Jira story has been moved to a sprint that is outside the current
        # PI and has therefore been removed from the current PI.
        metric_data.removed.append(planned_story)

    return None


# ==============================================================================
def update_pi_planned_detail_list(planned_detail_list: list[PlannedStoryDetailRec],
                                  planned_story: PlannedStoryRec,
                                  jira_story: FastStoryRec) -> None:

    if jira_story.sprints:
        cur_sprint = jira_story.sprints[0]
    else:
        cur_sprint = ''
    new_detail_rec = PlannedStoryDetailRec(planned_story.key,
                                           planned_story.summary,
                                           planned_story.points,
                                           jira_story.points,
                                           planned_story.sprint,
                                           cur_sprint,
                                           jira_story.status,
                                           planned_story.team)
    planned_detail_list.append(new_detail_rec)

    return None


# ==============================================================================
def process_unplanned_stories_added_to_pi(input_data: InputData, metrics_data: MetricsData) -> None:

    print('    Processing Unplanned Stories added to PI')
    for cur_jira_story in input_data.jira_stories:
        pi_planned_story = input_data.pi_planned_stories_data.get_pi_planned_story_data(cur_jira_story.issue_key)
        if pi_planned_story is None:
            # This jira story is not in the PI Planned stories and is therefore Unplanned
            # and added after PI had begun
            update_sprint_totals_for_unplanned_story(cur_jira_story, metrics_data)
            metrics_data.unplanned.append(cur_jira_story)
            # update_pi_unplanned_detail_list(cur_jira_story, metrics_data.unplanned)
        else:
            # This jira story was planned
            pass

    return None


# ==============================================================================
def update_sprint_totals_for_unplanned_story(jira_story: FastStoryRec, metric_data: MetricsData) -> None:

    # Check to see if the jira_story.sprints has been unassigned
    if jira_story.sprints:
        sprint_to_find = jira_story.sprints[0]
    else:
        sprint_to_find = 'Not Assigned to Sprint'
        metric_data.unassigned.append(jira_story)

    # Loop through the metrics_data list of sprints to find the correct sprint to
    # add this unplanned jira story points to the totals for that sprint in the PI
    found = False
    for cur_sprint in metric_data.sprints:
        if cur_sprint.sprint_name == sprint_to_find:
            found = True
            # Update the unplanned totals for the cur_sprint
            cur_sprint.unplanned_totals.stories += 1
            cur_sprint.unplanned_totals.points += jira_story.points
            # Update the planned + unplanned totals for the cur_sprint
            cur_sprint.combined_totals.stories += 1
            cur_sprint.combined_totals.points += jira_story.points
            # Check status of Jira story to see if the status is Done
            if jira_story.status == 'Done':
                # Update the Done unplanned totals for the cur_sprint
                cur_sprint.unplanned_totals.done_stories += 1
                cur_sprint.unplanned_totals.done_points += jira_story.points
                # Update the planned + unplanned totals for the cur_sprint
                cur_sprint.combined_totals.done_stories += 1
                cur_sprint.combined_totals.done_points += jira_story.points
            break  # updated the totals for the correct sprint, stop the loop
    if not found:
        # The Jira story has been moved to a sprint that is outside the current
        # PI and has therefore been removed from the current PI.
        print('     Unplanned story ' + jira_story.issue_key + ' has been moved to a sprint outside the PI')
        cur_sprint = metric_data.sprints[len(metric_data.sprints) - 1]  # last sprint is the Not Assigned to sprint
        # Update the unplanned totals for the cur_sprint
        cur_sprint.unplanned_totals.stories += 1
        cur_sprint.unplanned_totals.points += jira_story.points
        # Update the planned + unplanned totals for the cur_sprint
        cur_sprint.combined_totals.stories += 1
        cur_sprint.combined_totals.points += jira_story.points
        # Check status of Jira story to see if the status is Done
        if jira_story.status == 'Done':
            # Update the Done unplanned totals for the cur_sprint
            cur_sprint.unplanned_totals.done_stories += 1
            cur_sprint.unplanned_totals.done_points += jira_story.points
            # Update the planned + unplanned totals for the cur_sprint
            cur_sprint.combined_totals.done_stories += 1
            cur_sprint.combined_totals.done_points += jira_story.points

    return None


# ==============================================================================
def create_pi_planning_metrics_report_ss(metrics_data: MetricsData, input_data: InputData) -> None:
    print('\n  Creating PI Planning Metrics spreadsheet')

    # create the spreadsheet workbook
    workbook = xlsxwriter.Workbook('Output files/' + date.today().strftime("%y-%m-%d - ") + input_data.pi_info.pi_name
                                   + ' Status' + '.xlsx')
    cell_formats = create_cell_formatting_options(workbook)

    write_the_pi_planning_metrics_tab_to_wb(workbook, cell_formats, input_data.pi_info, metrics_data)

    write_planned_stories_detail_tab_to_wb(workbook, cell_formats, metrics_data.planned)

    write_unplanned_stories_detail_tab_to_wb(workbook, cell_formats, metrics_data.unplanned)

    write_jira_stories_detail_to_wb(workbook, cell_formats, input_data.jira_stories)

    write_no_sprint_assigned_tab_to_wb(workbook, cell_formats, metrics_data.unassigned)

    write_removed_stories_data_to_wb(workbook, cell_formats, metrics_data.removed)

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
def write_the_pi_planning_metrics_tab_to_wb(workbook, cell_fmts: Type[CellFormats], pi_info: ProgramIncrementRec,
                                            metrics_data: MetricsData) -> None:
    print('    Writing PI Planning Metrics worksheet')

    pi_planning_worksheet = workbook.add_worksheet('PI Planning Metrics')
    # Setup Column Width for the Columns in the worksheet A through D
    pi_planning_worksheet.set_column('B:B', 30, cell_fmts.metrics_ws_fmt)
    pi_planning_worksheet.set_column('C:G', 24, cell_fmts.metrics_ws_fmt)

    write_program_increment_elapsed_days_info_to_metrics_ws(pi_planning_worksheet, cell_fmts, pi_info)

    write_planned_sprints_data_to_metrics_ws(pi_planning_worksheet, cell_fmts, metrics_data.sprints)

    write_the_planned_story_confidence_level_data_to_metrics_ws(pi_planning_worksheet, cell_fmts, metrics_data.sprints)

    write_unplanned_sprints_data_to_metrics_ws(pi_planning_worksheet, cell_fmts, metrics_data.sprints)

    write_the_unplanned_story_confidence_level_data_to_metrics_ws(pi_planning_worksheet, cell_fmts, metrics_data.sprints)

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
def write_program_increment_elapsed_days_info_to_metrics_ws(worksheet, cell_fmts, pi_info: ProgramIncrementRec) -> None:

    print('    Writing PI Workdays Elapsed table to Metrics worksheet')

    # build the list of holidays for 2023
    first_day_in_pi = pi_info.pi_start_date
    last_day_in_pi = pi_info.pi_end_date
    holidays = [datetime(2023, 1, 2), datetime(2023, 2, 20), datetime(2023, 5, 29), datetime(2023, 7, 4),
                datetime(2023, 11, 23), datetime(2023, 11, 24), datetime(2023, 12, 25)]
    # Get the total number of workdays in the Program Increment
    total_pi_workdays = networkdays(first_day_in_pi, last_day_in_pi, holidays)

    # Get the total number of workdays that have already elapsed in the Program Increment
    elapsed_pi_workdays = networkdays(first_day_in_pi, datetime.now(), holidays)

    # Calculate the percentage of the Program Increment that has elapsed
    pi_elapsed = elapsed_pi_workdays / total_pi_workdays

    # Create the table data
    table_data = [
        [pi_info.pi_name,
         pi_info.pi_start_date.strftime("%m/%d/%Y"),
         pi_info.pi_end_date.strftime("%m/%d/%Y"),
         total_pi_workdays,
         elapsed_pi_workdays,
         pi_elapsed]
    ]

    # ******************************************************************
    # Set Status Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    pi_elapsed_table = calc_table_starting_and_ending_cells(2, 'B', 'G', 1, False)

    worksheet.add_table(pi_elapsed_table,
                        {'name': 'PI_Elapsed_Table',
                         'style': 'Table Style Medium 2',
                         'autofilter': False,
                         'first_column': True,
                         'data': table_data,
                         'columns': [
                             {'header': 'Program Increment', 'header_format': cell_fmts.header_left_fmt,
                              'format': cell_fmts.left_fmt},
                             {'header': 'PI Start Date', 'header_format': cell_fmts.header_center_fmt,
                              'format': cell_fmts.center_fmt},
                             {'header': 'PI End Date', 'header_format': cell_fmts.header_center_fmt,
                              'format': cell_fmts.center_fmt},
                             {'header': '# Workdays Total', 'header_format': cell_fmts.header_center_fmt,
                              'format': cell_fmts.center_fmt},
                             {'header': '# Workdays Elapsed', 'header_format': cell_fmts.header_center_fmt,
                              'format': cell_fmts.center_fmt},
                             {'header': '% of Workdays Elapsed', 'header_format': cell_fmts.header_center_fmt,
                              'format': cell_fmts.percent_center_fmt}]
                         })
    worksheet.write('B1', 'PI Workdays Info', cell_fmts.table_label_fmt)

    return None


# ==============================================================================
def write_planned_sprints_data_to_metrics_ws(worksheet, cell_fmts, sprints: list[SprintRec]) -> None:

    print('    Writing Planned Sprints Totals table to Metrics worksheet')

    table_data = []
    for cur_sprint in sprints:
        new_row_data = [cur_sprint.sprint_name,
                        cur_sprint.planned_totals.stories,
                        cur_sprint.planned_totals.points,
                        cur_sprint.planned_totals.done_points]
        table_data.append(new_row_data)

    # ******************************************************************
    # Set Sprints Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    sprints_table = calc_table_starting_and_ending_cells(6, 'B', 'E', len(sprints), True)
    worksheet.add_table(sprints_table,
                        {'name': 'Sprints_Table',
                         'style': 'Table Style Medium 2',
                         'autofilter': False,
                         'first_column': True,
                         'total_row': 1,
                         'data': table_data,
                         'columns': [
                             {'header': 'Sprints', 'header_format': cell_fmts.header_left_fmt,
                              'total_string': 'Grand Total', 'format': cell_fmts.left_fmt},
                             {'header': 'Total # Stories', 'total_function': 'sum', 'format': cell_fmts.right_fmt},
                             {'header': 'Total # Points', 'total_function': 'sum', 'format': cell_fmts.right_fmt},
                             {'header': 'Completed Points', 'total_function': 'sum', 'format': cell_fmts.right_fmt}]
                         })
    worksheet.write('B5', 'Planned PI Story Totals', cell_fmts.table_label_fmt)
    return None


# ==============================================================================
def write_the_planned_story_confidence_level_data_to_metrics_ws(worksheet, cell_fmts, sprints: list[SprintRec]) -> None:

    print('    Writing Unplanned Confidence Level table to Metrics worksheet')

    # loop through the sprints for planned_totals and total up the counts, points, and done points
    count = 0
    points = 0
    done_count = 0
    done_points = 0
    done_percentage = 0.0
    for cur_sprint in sprints:
        count += cur_sprint.planned_totals.stories
        points += cur_sprint.planned_totals.points
        done_count += cur_sprint.planned_totals.done_stories
        done_points += cur_sprint.planned_totals.done_points
        if points != 0:
            done_percentage = done_points / points

    # Calculate the 80% Planned Confidence Level counts, points, and done points
    pcl_count = round(count * .8)
    pcl_points = round(points * .8)
    pcl_done_percentage = pcl_points / points

    table_data = [
        ['Confidence Level Goal', pcl_count, pcl_points, pcl_done_percentage],
        ['Actual Completion to Date', done_count, done_points, done_percentage]
    ]

    # ******************************************************************
    # Set Status Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    confidence_tbl = calc_table_starting_and_ending_cells(18, 'B', 'E', 2, False)

    worksheet.add_table(confidence_tbl,
                        {'name': 'Confidence_Table',
                         'style': 'Table Style Medium 2',
                         'autofilter': False,
                         'first_column': True,
                         'total_row': 0,
                         'data': table_data,
                         'columns': [{'header': 'Results', 'header_format': cell_fmts.header_left_fmt,
                                      'format': cell_fmts.left_fmt},
                                     {'header': 'Stories', 'format': cell_fmts.right_fmt},
                                     {'header': 'Points', 'format': cell_fmts.right_fmt},
                                     {'header': 'Completion Percentage', 'format': cell_fmts.percent_center_fmt}
                                     ]
                         }
                        )
    worksheet.write('B17', 'Planned Confidence Level of Success', cell_fmts.table_label_fmt)
    return None


# ==============================================================================
def write_unplanned_sprints_data_to_metrics_ws(worksheet, cell_fmts, sprints: list[SprintRec]) -> None:

    print('    Writing Unplanned Sprints Totals table to Metrics worksheet')

    table_data = []
    for cur_sprint in sprints:
        new_row_data = [cur_sprint.sprint_name,
                        cur_sprint.unplanned_totals.stories,
                        cur_sprint.unplanned_totals.points,
                        cur_sprint.unplanned_totals.done_points]
        table_data.append(new_row_data)

    # ******************************************************************
    # Set Sprints Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    unplanned_table = calc_table_starting_and_ending_cells(24, 'B', 'E', len(sprints), True)
    worksheet.add_table(unplanned_table,
                        {'name': 'Unplanned_Table',
                         'style': 'Table Style Medium 2',
                         'autofilter': False,
                         'first_column': True,
                         'total_row': 1,
                         'data': table_data,
                         'columns': [
                             {'header': 'Sprints', 'header_format': cell_fmts.header_left_fmt,
                              'total_string': 'Grand Total', 'format': cell_fmts.left_fmt},
                             {'header': 'Stories', 'total_function': 'sum', 'format': cell_fmts.right_fmt},
                             {'header': 'Points', 'total_function': 'sum', 'format': cell_fmts.right_fmt},
                             {'header': 'Completed Points', 'total_function': 'sum', 'format': cell_fmts.right_fmt}]
                         })
    worksheet.write('B23', 'Unplanned PI Story Totals', cell_fmts.table_label_fmt)
    return None


# ==============================================================================
def write_the_unplanned_story_confidence_level_data_to_metrics_ws(worksheet, cell_fmts,
                                                                  sprints: list[SprintRec]) -> None:

    print('    Writing Planned + Unplanned Confidence Level table to Metrics worksheet')

    # loop through the sprints for unplanned_totals and total up the counts, points, and done points
    count = 0
    points = 0
    done_count = 0
    done_points = 0
    done_percentage = 0.0
    for cur_sprint in sprints:
        count += cur_sprint.combined_totals.stories
        points += cur_sprint.combined_totals.points
        done_count += cur_sprint.combined_totals.done_stories
        done_points += cur_sprint.combined_totals.done_points
        if points != 0:
            done_percentage = done_points / points

    # Calculate the 80% Planned + Unplanned Confidence Level counts, points, and done points
    cl_count = round(count * .8)
    cl_points = round(points * .8)
    cl_done_percentage = cl_points / points

    table_data = [
        ['Confidence Level Goal', cl_count, cl_points, cl_done_percentage],
        ['Actual Completion to Date', done_count, done_points, done_percentage]
    ]

    # ******************************************************************
    # Set Status Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, Total_row)
    # ******************************************************************
    confidence_tbl_unplanned = calc_table_starting_and_ending_cells(36, 'B', 'E', 2, False)

    worksheet.add_table(confidence_tbl_unplanned,
                        {'name': 'Confidence_Table_2',
                         'style': 'Table Style Medium 2',
                         'autofilter': False,
                         'first_column': True,
                         'total_row': 0,
                         'data': table_data,
                         'columns': [{'header': 'Results', 'header_format': cell_fmts.header_left_fmt,
                                      'format': cell_fmts.left_fmt},
                                     {'header': 'Stories', 'format': cell_fmts.right_fmt},
                                     {'header': 'Points', 'format': cell_fmts.right_fmt},
                                     {'header': 'Completion Percentage', 'format': cell_fmts.percent_center_fmt}
                                     ]
                         }
                        )
    worksheet.write('B35', 'Planned + Unplanned Confidence Level of Success', cell_fmts.table_label_fmt)
    return None


# ==============================================================================
def write_planned_stories_detail_tab_to_wb(workbook, cell_fmts, planned_stories: list[PlannedStoryDetailRec]) -> None:
    print('    Writing Planned Story Detail worksheet tab')
    planned_story_detail_ws = workbook.add_worksheet('Planned Stories Detail')

    # Setup table layout
    planned_story_detail_ws.set_column('A:A', 12, cell_fmts.left_fmt)
    planned_story_detail_ws.set_column('B:B', 80, cell_fmts.left_fmt)
    planned_story_detail_ws.set_column('C:F', 18, cell_fmts.center_fmt)
    planned_story_detail_ws.set_column('G:H', 18, cell_fmts.center_fmt)

    # ******************************************************************
    # Set Excel Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows, total_row)
    # ******************************************************************
    num_planned_stories = len(planned_stories)
    planned_stories_tbl = calc_table_starting_and_ending_cells(1, 'A', 'H', num_planned_stories, False)
    if num_planned_stories > 0:
        table_data = []
        for cur_story in planned_stories:
            new_row_data = [cur_story.key,
                            cur_story.summary,
                            cur_story.plan_points,
                            cur_story.cur_points,
                            cur_story.plan_sprint,
                            cur_story.cur_sprint,
                            cur_story.status,
                            cur_story.team]
            table_data.append(new_row_data)

        planned_story_detail_ws.add_table(planned_stories_tbl,
                                          {'name': 'Planned_Stories_table',
                                           'style': 'Table Style Medium 2',
                                           'autofilter': True,
                                           'first_column': False,
                                           'data': table_data,
                                           'columns': [{'header': 'Issue Key', 'format': cell_fmts.left_fmt},
                                                       {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                                       {'header': 'Planned Points', 'format': cell_fmts.center_fmt},
                                                       {'header': 'Current Points', 'format': cell_fmts.center_fmt},
                                                       {'header': 'Planned Sprint', 'format': cell_fmts.left_fmt},
                                                       {'header': 'Current Sprint', 'format': cell_fmts.left_fmt},
                                                       {'header': 'Current Status', 'format': cell_fmts.left_fmt},
                                                       {'header': 'Team', 'format': cell_fmts.left_fmt}]
                                           }
                                          )

    return None


# ==============================================================================
def write_jira_stories_detail_to_wb(workbook, cell_fmts: Type[CellFormats], jira_stories: FastStoryData) -> None:

    print('    Writing Jira Stories Detail worksheet tab')
    jira_data_ws = workbook.add_worksheet('Jira Stories Detail')
    # Setup Jira table layout
    jira_data_ws.set_column('A:B', 12, cell_fmts.center_fmt)
    jira_data_ws.set_column('C:C', 80, cell_fmts.center_fmt)
    jira_data_ws.set_column('D:D', 20, cell_fmts.center_fmt)
    jira_data_ws.set_column('E:E', 22, cell_fmts.center_fmt)
    jira_data_ws.set_column('F:F', 12, cell_fmts.center_fmt)
    jira_data_ws.set_column('G:G', 12, cell_fmts.center_fmt)
    jira_data_ws.set_column('H:I', 24, cell_fmts.center_fmt)
    jira_data_ws.set_column('J:J', 20, cell_fmts.center_fmt)

    # ******************************************************************
    # Set Jira Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    jira_data_tbl = calc_table_starting_and_ending_cells(1, 'A', 'J', len(jira_stories.stories), False)
    table_data = []
    for cur_story in jira_stories.stories:
        if cur_story.sprints:
            cur_sprint = cur_story.sprints[0]  # First sprint in list is the current sprint
        else:
            cur_sprint = ''
        if cur_story.labels:
            cur_labels = '; '.join(cur_story.labels)  # Convert list to string
        else:
            cur_labels = ''
        new_row_data = [cur_story.issue_key,
                        cur_story.issue_type,
                        cur_story.summary,
                        cur_story.status,
                        cur_story.assignee,
                        cur_story.points,
                        cur_labels,
                        cur_story.created.strftime("%m/%d/%Y, %H:%M:%S"),
                        cur_sprint]
        table_data.append(new_row_data)

    jira_data_ws.add_table(jira_data_tbl,
                           {'name': 'Jira_Data_table',
                            'style': 'Table Style Medium 2',
                            'autofilter': True,
                            'first_column': False,
                            'data': table_data,
                            'columns': [{'header': 'Issue Key', 'format': cell_fmts.left_fmt},
                                        {'header': 'Issue Type', 'format': cell_fmts.center_fmt},
                                        {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                        {'header': 'Status', 'format': cell_fmts.left_fmt},
                                        {'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                        {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                        {'header': 'Labels', 'format': cell_fmts.left_fmt},
                                        {'header': 'Created', 'format': cell_fmts.left_fmt},
                                        {'header': 'Sprint', 'format': cell_fmts.left_fmt}]
                            })

    return None


# ==============================================================================
def write_no_sprint_assigned_tab_to_wb(workbook, cell_fmts, unassigned_jira_stories: list[PiPlanStoryRec]) -> None:
    print('    Writing Unassigned Stories worksheet tab')
    unassigned_worksheet = workbook.add_worksheet('No Sprint Assigned')

    # Setup Details table layout
    unassigned_worksheet.set_column('A:B', 12, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('C:C', 80, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('D:D', 20, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('E:E', 22, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('F:F', 12, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('G:G', 12, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('H:I', 24, cell_fmts.center_fmt)
    unassigned_worksheet.set_column('J:J', 20, cell_fmts.center_fmt)

    # ******************************************************************
    # Set Jira Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    unassigned_data_tbl = calc_table_starting_and_ending_cells(1, 'A', 'J', len(unassigned_jira_stories), False)
    if len(unassigned_jira_stories) > 0:
        table_data = []
        for cur_story in unassigned_jira_stories:
            new_row_data = [cur_story.issue_key,
                            cur_story.issue_type,
                            cur_story.summary,
                            cur_story.status,
                            cur_story.assignee,
                            cur_story.points,
                            ';'.join(cur_story.labels),
                            cur_story.created.strftime("%m/%d/%Y, %H:%M:%S"),
                            ';'.join(cur_story.sprints)]
            table_data.append(new_row_data)

        unassigned_worksheet.add_table(unassigned_data_tbl,
                                       {'name': 'Unassigned_Stories_table',
                                        'style': 'Table Style Medium 2',
                                        'autofilter': True,
                                        'first_column': False,
                                        'data': table_data,
                                        'columns': [{'header': 'Issue Key', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Issue Type', 'format': cell_fmts.center_fmt},
                                                    {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Status', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                                    {'header': 'Labels', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Created', 'format': cell_fmts.left_fmt},
                                                    {'header': 'Sprint', 'format': cell_fmts.left_fmt}]
                                        }
                                       )
    else:
        unassigned_worksheet.write('B2', 'No Unassigned Stories to Report', cell_fmts.left_bold_fmt)

    return None


# ==============================================================================
def write_removed_stories_data_to_wb(workbook, cell_fmts, stories_removed_from_pi: list[PlannedStoryRec]) -> None:
    print('    Writing Removed Stories worksheet tab')
    removed_worksheet = workbook.add_worksheet('Planned - Removed from PI')

    # Setup table layout
    removed_worksheet.set_column('A:A', 12, cell_fmts.center_fmt)
    removed_worksheet.set_column('B:B', 80, cell_fmts.center_fmt)
    removed_worksheet.set_column('C:E', 14, cell_fmts.center_fmt)

    # ******************************************************************
    # Set Jira Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    removed_stories_tbl = calc_table_starting_and_ending_cells(1, 'A', 'E', len(stories_removed_from_pi), False)
    if len(stories_removed_from_pi) > 0:
        table_data = []
        for cur_story in stories_removed_from_pi:
            new_row_data = [cur_story.key,
                            cur_story.summary,
                            cur_story.points,
                            cur_story.team,
                            cur_story.sprint]
            table_data.append(new_row_data)

        removed_worksheet.add_table(removed_stories_tbl,
                                    {'name': 'Removed_Stories_table',
                                     'style': 'Table Style Medium 2',
                                     'autofilter': True,
                                     'first_column': False,
                                     'data': table_data,
                                     'columns': [{'header': 'Issue Key', 'format': cell_fmts.left_fmt},
                                                 {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                                 {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                                 {'header': 'Team', 'format': cell_fmts.left_fmt},
                                                 {'header': 'Sprint', 'format': cell_fmts.left_fmt}]
                                     }
                                    )
    else:
        removed_worksheet.write('B2', 'No Removed Stories to Report', cell_fmts.left_bold_fmt)

    return None


# ==============================================================================
def write_unplanned_stories_detail_tab_to_wb(workbook, cell_fmts, unplanned_jira_stories: list[PiPlanStoryRec]) -> None:
    print('    Writing Unplanned Stories Detail worksheet tab')
    unplanned_worksheet = workbook.add_worksheet('Unplanned Stories Detail')

    # Setup Details table layout
    unplanned_worksheet.set_column('A:B', 12, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('C:C', 80, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('D:D', 20, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('E:E', 22, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('F:F', 12, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('G:G', 12, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('H:I', 24, cell_fmts.center_fmt)
    unplanned_worksheet.set_column('J:J', 20, cell_fmts.center_fmt)

    # ******************************************************************
    # Set Jira Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    unplanned_data_tbl = calc_table_starting_and_ending_cells(1, 'A', 'J', len(unplanned_jira_stories), False)
    if len(unplanned_jira_stories) > 0:
        table_data = []
        for cur_story in unplanned_jira_stories:
            if cur_story.sprints:
                cur_sprint = cur_story.sprints[0]  # First sprint in list is the current sprint
            else:
                cur_sprint = ''
            if cur_story.labels:
                cur_labels = '; '.join(cur_story.labels)  # Convert list to string
            else:
                cur_labels = ''
            new_row_data = [cur_story.issue_key,
                            cur_story.issue_type,
                            cur_story.summary,
                            cur_story.status,
                            cur_story.assignee,
                            cur_story.points,
                            cur_labels,
                            cur_story.created.strftime("%m/%d/%Y, %H:%M:%S"),
                            cur_sprint]
            table_data.append(new_row_data)

        unplanned_worksheet.add_table(unplanned_data_tbl,
                                      {'name': 'Unplanned_Stories_table',
                                       'style': 'Table Style Medium 2',
                                       'autofilter': True,
                                       'first_column': False,
                                       'data': table_data,
                                       'columns': [{'header': 'Issue Key', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Issue Type', 'format': cell_fmts.center_fmt},
                                                   {'header': 'Summary', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Status', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                                   {'header': 'Labels', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Created', 'format': cell_fmts.left_fmt},
                                                   {'header': 'Sprint', 'format': cell_fmts.left_fmt}]
                                       }
                                      )
    else:
        unplanned_worksheet.write('B2', 'No Unplanned Stories to Report', cell_fmts.left_bold_fmt)

    return None


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************

def create_pi_planning_metrics():
    print('\nBegin PI Planning Metrics')
    input_data = get_input_file_data()
    if input_data.success:
        metrics_data = process_program_increment_data(input_data)
        if metrics_data.success:
            create_pi_planning_metrics_report_ss(metrics_data, input_data)
        else:
            print('Terminating Error Processing PI Data')
    else:
        print('Terminating - Error getting Input Data')
    print('End PI Planning Metrics')

    return None


if __name__ == "__main__":
    create_pi_planning_metrics()
