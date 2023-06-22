#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime, date

# Third party imports
from typing import Type
import xlsxwriter

# local application imports

# SGM Shared Module imports
from kclFastSharedDataClasses import *
from kclGetCsvLettersData import GetCsvLetterData, CsLetterRec
from kclGetFastStoryDataJiraAPI import FastStoryData, FastStoryRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************

@dataclass()
class InputData:
    jql_query: str = ''
    report_type: str = ''
    output_filename: str = ''
    letter_info: list[CsLetterRec] = None
    sprint_to_process: str = ''
    team_info: list[TeamRec] = None
    jira_stories: FastStoryData = None
    success: bool = False


@dataclass
class LetterStoryRec:
    letter_id: str = ''
    issue_key: str = ''
    summary: str = ''
    status: str = ''
    assignee: str = ''
    test_assignee: str = ''
    points: int = 0
    sprints: str = ''
    labels: str = ''
    priority: str = ''
    issue_type: str = ''
    created: datetime = datetime(1900, 1, 1, 00, 00, 00, 342380)
    is_blocked_by: str = ''
    blocks: str = ''


@dataclass
class LetterData:
    letter_id: str = ''
    description: str = ''
    points_total: int = 0
    points_done: int = 0
    contains_story_in_cur_sprint: bool = False
    jira_stories: list[LetterStoryRec] = field(default_factory=list)


@dataclass
class CellFormats:
    metrics_ws_fmt = None
    left_fmt = None
    left_green_fmt = None
    left_red_fmt = None
    left_orange_fmt = None
    left_bold_fmt = None
    left_bold_separator_fmt = None
    left_done_separator_fmt = None
    header_left_fmt = None
    header_center_fmt = None
    right_fmt = None
    percent_fmt = None
    percent_right_fmt = None
    center_fmt = None
    def_fmt = None
    table_label_fmt = None
    total_row_fmt = None
    grand_total_row_fmt = None
    grand_percent_right_fmt = None


# ******************************************************************************
# ******************************************************************************
# # * Functions
# ******************************************************************************
# ******************************************************************************

# ==============================================================================
def get_input_data():
    print('\n  Begin Getting Input Data ')

    input_data = InputData()

    # get the type of letter to report on, CS Letter or Claim Letter
    get_letter_report_to_create(input_data)
    # get the Sprint to process from the user via console input
    input_data.sprint_to_process = get_sprint_to_process()
    # Get Get Letters Data, Letter ID, Letter Description from the FastSprintInfo.csv spreadsheet
    print('\n  Getting Letter Codes from CSV file')
    input_data.letter_info = GetCsvLetterData(Path.cwd())
    if input_data.letter_info is not None:
        # input_data.sprint_info = fast_sprint_info.get_sprint_info(sprint_to_process)
        # Get the FAST Jira Story data for the sprint being processed
        input_data.jira_stories = FastStoryData(input_data.jql_query).stories
        if input_data.jira_stories is not None:
            input_data.success = True
            print('   Success Getting Input Data')
        else:
            print('   *** Error getting Sprint Info from SGM - Jira - FAST Sprint Data (Jira).csv')
    else:
        print('   *** Error getting Sprint Info from FastSprintInfo.csv')

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
def get_report_to_launch() -> int:
    app_to_launch: int = 0
    valid_input: bool = False

    while not valid_input:
        print('\n')
        print('   ******************************************************************')
        print('   **                                                             ***')
        print('   **    Select the Report to Create                              ***')
        print('   **                                                             ***')
        print('   **         1 - Create FAST CS Letters Report                   ***')
        print('   **         2 - Create FAST Claim Letters Report                ***')
        print('   **         3 - Correspondence Prod Issue Report                ***')
        print('   **                                                             ***')
        print('   ******************************************************************')

        user_input = input('\n   Enter Number to Create  ==> ')
        match user_input:
            case '1':
                app_to_launch = int(user_input)
                valid_input = True
            case '2':
                app_to_launch = int(user_input)
                valid_input = True
            case '3':
                app_to_launch = int(user_input)
                valid_input = True
            case _:
                valid_input = False
                print('\n\n\n   Invalid Report Number, valid Report Numbers are 1 & 2')

    return app_to_launch


# ==============================================================================
def get_letter_report_to_create(input_data: InputData) -> None:
    done = False
    while not done:
        report_to_create = get_report_to_launch()
        match report_to_create:
            case 1:
                input_data.jql_query = CS_LETTER_DATA_JQL
                input_data.report_type = 'CS_Letters'
                input_data.output_filename = 'CS Letter Report.xlsx'
                done = True
            case 2:
                input_data.jql_query = CLAIM_LETTER_DATA_JQL
                input_data.report_type = 'Claim_Letters'
                input_data.output_filename = 'Claim Letter Report.xlsx'
                done = True
            case 3:
                input_data.jql_query = CORRESPONDENCE_PROD_ISSUE
                input_data.report_type = 'Correspondence_Letters'
                input_data.output_filename = 'Correspondence Letter Prod Issue Report.xlsx'
                done = True

    return None


# ==============================================================================
def process_cs_letter_data(input_data: InputData) -> list[LetterData]:
    print('\n   Begin Processing CS Letter Data')
    letter_data: list[LetterData] = []
    for cur_jira_story in input_data.jira_stories:
        new_letter_story = create_letter_story_from_jira_story(cur_jira_story, input_data.report_type)
        letter_type_found = False
        if letter_data is not None:
            for cur_letter_type in letter_data:
                if new_letter_story.letter_id == cur_letter_type.letter_id:
                    cur_letter_type.jira_stories.append(new_letter_story)
                    cur_letter_type.points_total += new_letter_story.points
                    if cur_jira_story.sprints:
                        if input_data.sprint_to_process in cur_jira_story.sprints:
                            cur_letter_type.contains_story_in_cur_sprint = True
                    if new_letter_story.status == 'Done':
                        cur_letter_type.points_done += new_letter_story.points
                    letter_type_found = True
                    break
        if letter_type_found:
            pass
        else:
            description = get_description_for_letter_type(new_letter_story.letter_id, input_data.letter_info)
            new_letter_type = LetterData(new_letter_story.letter_id, description)
            new_letter_type.points_total = new_letter_story.points
            if cur_jira_story.sprints:
                if input_data.sprint_to_process in cur_jira_story.sprints:
                    new_letter_type.contains_story_in_cur_sprint = True
            if new_letter_story.status == 'Done':
                new_letter_type.points_done = new_letter_story.points
            new_letter_type.jira_stories.append(new_letter_story)
            letter_data.append(new_letter_type)

    return letter_data


# ==============================================================================
def get_description_for_letter_type(letter_id: str, letter_info: GetCsvLetterData) -> str:
    description = letter_info.get_letter_description(letter_id)
    if description == '':
        description = 'Unknown Letter ID'

    return description


# ==============================================================================
def create_letter_story_from_jira_story(jira_story: FastStoryRec, report_type) -> LetterStoryRec:
    new_rec = LetterStoryRec()
    if report_type == 'Correspondence_Letters':
        new_rec.letter_id = 'Prod Issue'
    else:
        letter_id_start: int = jira_story.summary.find('(') + 1
        letter_id_end: int = jira_story.summary.find(')')
        new_rec.letter_id = jira_story.summary[letter_id_start: letter_id_end]
    new_rec.issue_key = jira_story.issue_key
    new_rec.summary = jira_story.summary
    new_rec.status = jira_story.status
    new_rec.assignee = jira_story.assignee
    new_rec.test_assignee = jira_story.test_assignee
    new_rec.points = jira_story.points
    new_rec.sprints = jira_story.sprints
    new_rec.labels = jira_story.labels
    new_rec.priority = jira_story.priority
    new_rec.issue_type = jira_story.issue_type
    new_rec.created = jira_story.created
    new_rec.is_blocked_by = jira_story.is_blocked_by

    return new_rec


# ==============================================================================
def create_letter_report_spreadsheet(letter_data: list[LetterData], report_filename: str, sprint_name: str) -> None:
    print('\n   Creating CS Letter Report spreadsheet')

    # create the spreadsheet workbook
    relative_path = 'Output files/' + date.today().strftime("%y-%m-%d") + ' ' + report_filename
    workbook = xlsxwriter.Workbook(relative_path)
    # create the cell formatting options for the workbook
    cell_fmts = create_cell_formatting_options(workbook)

    #  Add the CS Letters worksheet to the workbook
    letters_ws = workbook.add_worksheet('Letters Detail')
    letters_ws.freeze_panes(1, 0)
    # Set the column widths and default cell formatting for the CS Letters worksheet
    create_cs_letter_stories_column_layout(letters_ws, cell_fmts)

    #  Add the Current Sprint Letters worksheet to the workbook
    cur_sprint_ws = workbook.add_worksheet('Current Sprint Letters')
    cur_sprint_ws.freeze_panes(1, 0)
    # Set the column widths and default cell formatting for the CS Letters worksheet
    create_cs_letter_stories_column_layout(cur_sprint_ws, cell_fmts)

    # Sort the letter data by letter ID so that the report shows status in alphabetical order to make
    # it easier to find specific letters on the report
    letter_data.sort(key=lambda letter_data_rec: letter_data_rec.letter_id)

    # Write the CS Letters Detail worksheet
    next_row_all_letters = write_header_row(letters_ws, cell_fmts)
    for cur_letter_type in letter_data:
        next_row_all_letters = write_the_letter_description_row_to_ws(cur_letter_type, next_row_all_letters, letters_ws,
                                                                      cell_fmts)
        next_row_all_letters = write_the_cs_letter_stories_for_description_to_ws(cur_letter_type.jira_stories,
                                                                                 sprint_name, next_row_all_letters,
                                                                                 letters_ws, cell_fmts)
        # Write the totals for the cur_letter_type
        next_row_all_letters = write_the_cur_letter_type_totals_row(cur_letter_type, next_row_all_letters, letters_ws, cell_fmts)
    write_the_cur_letter_grand_totals_row(letter_data, next_row_all_letters, letters_ws, cell_fmts)

    # Write the Current Sprint Letters detail worksheet
    next_row_all_letters = write_header_row(cur_sprint_ws, cell_fmts)
    for cur_letter_type in letter_data:
        if cur_letter_type.contains_story_in_cur_sprint:
            next_row_all_letters = write_the_letter_description_row_to_ws(cur_letter_type, next_row_all_letters, cur_sprint_ws, cell_fmts)
            next_row_all_letters = write_the_cs_letter_stories_for_description_to_ws(cur_letter_type.jira_stories, sprint_name,
                                                                         next_row_all_letters, cur_sprint_ws, cell_fmts)
            # Write the totals for the cur_letter_type
            next_row_all_letters = write_the_cur_letter_type_totals_row(cur_letter_type, next_row_all_letters, cur_sprint_ws, cell_fmts)
    # write_the_cur_letter_grand_totals_row(letter_data, next_row_all_letters, letters_ws, cell_fmts)

    workbook.close()
    print('   Completed Letter Report Spreadsheet')

    return None


# ==============================================================================
def create_cell_formatting_options(workbook) -> Type[CellFormats]:
    # create predefined cell_formats to be used for cells in the workbook
    cell_fmt = CellFormats
    cell_fmt.metrics_ws_fmt = workbook.add_format({'font_name': 'Calibri', 'align': 'center', 'font_size': 12})
    cell_fmt.left_fmt = workbook.add_format({'align': 'left', 'indent': 1})
    cell_fmt.left_green_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'font_color': 'green'})
    cell_fmt.left_red_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'font_color': 'red', 'bold': 1})
    cell_fmt.left_orange_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'font_color': 'orange', 'bold': 1})
    cell_fmt.left_bold_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'indent': 1})
    cell_fmt.left_bold_separator_fmt = workbook.add_format(
        {'align': 'left', 'bold': 1, 'indent': 1, 'bg_color': '#DA9694'})
    cell_fmt.left_done_separator_fmt = workbook.add_format(
        {'align': 'left', 'bold': 1, 'indent': 1, 'bg_color': '#C4D79B'})
    cell_fmt.header_left_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'indent': 1, 'font_size': 12})
    cell_fmt.header_center_fmt = workbook.add_format(
        {'align': 'center', 'bold': 1, 'font_size': 12, 'bg_color': '#B8CCE4'})
    cell_fmt.right_fmt = workbook.add_format({'align': 'right'})
    cell_fmt.percent_fmt = workbook.add_format({'align': 'right', 'indent': 8, 'num_format': '0%'})
    cell_fmt.percent_right_fmt = workbook.add_format({'align': 'right', 'bold': 1, 'num_format': '0%', 'top': 6})
    cell_fmt.center_fmt = workbook.add_format({'align': 'center'})
    cell_fmt.center_bold_fmt = workbook.add_format({'align': 'center', 'bold': 1})
    cell_fmt.def_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'text_wrap': 1})
    cell_fmt.table_label_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'font_size': 14})
    cell_fmt.total_row_fmt = workbook.add_format(({'align': 'right', 'bold': 1, 'top': 6}))
    cell_fmt.grand_total_row_fmt = workbook.add_format(
        ({'align': 'right', 'bold': 1, 'top': 6, 'bg_color': '#B8CCE4', 'font_size': 12}))
    cell_fmt.grand_percent_right_fmt = workbook.add_format(
        {'align': 'right', 'num_format': '0%', 'top': 6, 'bg_color': '#B8CCE4', 'font_size': 12})

    return cell_fmt


# ==============================================================================
def create_cs_letter_stories_column_layout(jira_stories_ws, cell_fmts: Type[CellFormats]) -> None:
    # Setup Jira table layout
    jira_stories_ws.set_column('A:B', 12, cell_fmts.center_fmt)
    jira_stories_ws.set_column('C:C', 80, cell_fmts.center_fmt)
    jira_stories_ws.set_column('D:D', 25, cell_fmts.center_fmt)
    jira_stories_ws.set_column('E:F', 16, cell_fmts.center_fmt)
    jira_stories_ws.set_column('G:H', 20, cell_fmts.center_fmt)
    jira_stories_ws.set_column('I:J', 12, cell_fmts.center_fmt)

    return None


# ==============================================================================
def write_header_row(ws, cell_fmts: Type[CellFormats]) -> int:
    ws.write(0, 0, 'Letter ID', cell_fmts.header_center_fmt)
    ws.write(0, 1, 'Key', cell_fmts.header_center_fmt)
    ws.write(0, 2, 'Summary', cell_fmts.header_center_fmt)
    ws.write(0, 3, 'Status', cell_fmts.header_center_fmt)
    ws.write(0, 4, 'Sprint', cell_fmts.header_center_fmt)
    ws.write(0, 5, 'Is Blocked By', cell_fmts.header_center_fmt)
    ws.write(0, 6, 'Assignee', cell_fmts.header_center_fmt)
    ws.write(0, 7, 'Test Assignee', cell_fmts.header_center_fmt)
    ws.write(0, 8, 'Points', cell_fmts.header_center_fmt)
    ws.write(0, 9, 'Points Done', cell_fmts.header_center_fmt)
    ws.write(0, 10, '% Done', cell_fmts.header_center_fmt)
    next_row = 1

    return next_row


# ==============================================================================
def write_the_letter_description_row_to_ws(letter_type: LetterData, cur_row: int, ws, cell_fmts: Type[CellFormats]):
    if letter_type.points_done == letter_type.points_total:
        separator_fmt = cell_fmts.left_done_separator_fmt
    else:
        separator_fmt = cell_fmts.left_bold_separator_fmt
    ws.write(cur_row, 0, letter_type.letter_id, separator_fmt)
    ws.write(cur_row, 1, letter_type.description, separator_fmt)
    ws.write(cur_row, 2, '', separator_fmt)
    ws.write(cur_row, 3, '', separator_fmt)
    ws.write(cur_row, 4, '', separator_fmt)
    ws.write(cur_row, 5, '', separator_fmt)
    ws.write(cur_row, 6, '', separator_fmt)
    ws.write(cur_row, 7, '', separator_fmt)
    ws.write(cur_row, 8, '', separator_fmt)
    ws.write(cur_row, 9, '', separator_fmt)
    ws.write(cur_row, 10, '', separator_fmt)

    next_row = cur_row + 1

    return next_row


# ==============================================================================
def write_the_cs_letter_stories_for_description_to_ws(letter_stories: list[LetterStoryRec],
                                                      sprint_name: str,
                                                      cur_row: int,
                                                      ws,
                                                      cell_fmts: Type[CellFormats]) -> int:
    for cur_letter_story in letter_stories:
        ws.write(cur_row, 1, cur_letter_story.issue_key, cell_fmts.left_fmt)
        ws.write(cur_row, 2, cur_letter_story.summary, cell_fmts.left_fmt)
        status_fmt = determine_story_status_format(cur_letter_story.status, cell_fmts)
        ws.write(cur_row, 3, cur_letter_story.status, status_fmt)
        sprint_fmt = determine_story_sprint_format(cur_letter_story.sprints, sprint_name, cur_letter_story.status,
                                                   cell_fmts)
        if cur_letter_story.sprints:
            ws.write(cur_row, 4, cur_letter_story.sprints[0], sprint_fmt)
        blocks = ', '.join(cur_letter_story.is_blocked_by)
        ws.write(cur_row, 5, blocks, cell_fmts.left_fmt)
        ws.write(cur_row, 6, cur_letter_story.assignee, cell_fmts.left_fmt)
        ws.write(cur_row, 7, cur_letter_story.test_assignee, cell_fmts.left_fmt)
        ws.write(cur_row, 8, cur_letter_story.points, cell_fmts.right_fmt)
        if cur_letter_story.status == 'Done':
            ws.write(cur_row, 9, cur_letter_story.points, cell_fmts.right_fmt)
        cur_row += 1

    next_row = cur_row + 1

    return next_row


# ==============================================================================
def determine_story_status_format(story_status: str, cell_fmts: Type[CellFormats]) -> Type[CellFormats]:
    match story_status:
        case 'Done':
            status_fmt = cell_fmts.left_green_fmt
        case 'UAT':
            status_fmt = cell_fmts.left_orange_fmt
        case _:
            status_fmt = cell_fmts.left_red_fmt

    return status_fmt


# ==============================================================================
def determine_story_sprint_format(story_sprints: list,
                                  cur_sprint: str,
                                  story_status: str,
                                  cell_fmts: Type[CellFormats]) -> Type[CellFormats]:
    sprint_fmt = cell_fmts.left_fmt
    if len(story_sprints) > 0:
        if story_sprints[0] == cur_sprint:
            match story_status:
                case 'Done':
                    sprint_fmt = cell_fmts.left_green_fmt
                case 'UAT':
                    sprint_fmt = cell_fmts.left_orange_fmt
                case _:
                    sprint_fmt = cell_fmts.left_red_fmt

    return sprint_fmt


# ==============================================================================
def write_the_cur_letter_type_totals_row(letter_type: LetterData, cur_row: int, ws,
                                         cell_fmts: Type[CellFormats]) -> int:
    total_row = cur_row - 1
    ws.write(total_row, 1, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 2, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 3, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 4, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 5, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 6, '', cell_fmts.total_row_fmt)
    ws.write(total_row, 7, 'Totals', cell_fmts.total_row_fmt)
    ws.write(total_row, 8, letter_type.points_total, cell_fmts.total_row_fmt)
    ws.write(total_row, 9, letter_type.points_done, cell_fmts.total_row_fmt)
    if letter_type.points_total > 0:
        percent_done = letter_type.points_done / letter_type.points_total
    else:
        percent_done = 0.0
    ws.write(total_row, 10, percent_done, cell_fmts.percent_right_fmt)
    next_row = total_row + 2
    return next_row


# ==============================================================================
def write_the_cur_letter_grand_totals_row(letter_types: list[LetterData], cur_row: int, ws,
                                          cell_fmts: Type[CellFormats]) -> None:
    grand_total_points: int = 0
    grand_total_done: int = 0
    grand_total_percent_done: float = 0.0

    total_row = cur_row + 1
    for cur_letter_type in letter_types:
        grand_total_points += cur_letter_type.points_total
        grand_total_done += cur_letter_type.points_done
    ws.write(total_row, 0, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 1, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 2, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 3, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 4, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 5, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 6, '', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 7, 'Grand Totals', cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 8, grand_total_points, cell_fmts.grand_total_row_fmt)
    ws.write(total_row, 9, grand_total_done, cell_fmts.grand_total_row_fmt)
    if grand_total_points > 0:
        grand_total_percent_done = grand_total_done / grand_total_points
    else:
        grand_total_percent_done = 0.0
    ws.write(total_row, 10, grand_total_percent_done, cell_fmts.grand_percent_right_fmt)


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************
def create_cs_letter_report():
    print('\nBegin Create CS Letter Report')
    input_data = get_input_data()
    if input_data.success:
        letter_data = process_cs_letter_data(input_data)
        if len(letter_data) > 0:
            create_letter_report_spreadsheet(letter_data, input_data.output_filename, input_data.sprint_to_process)

    print('\nEnd Create CS Letter Report')

    return None


if __name__ == "__main__":
    create_cs_letter_report()
