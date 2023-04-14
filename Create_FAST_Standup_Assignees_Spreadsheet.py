#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************
# Standard library imports
import sys
from pathlib import Path
from dataclasses import dataclass
from typing import Type

# Third party imports
import xlsxwriter

# local file imports
from Create_
# SGM Shared Module imports
sys.path.append('C:/Users/kap3309/OneDrive - Kansas City Life Insurance/PythonDev/Modules')
from kclGetCsvJiraStoryData import CsvJiraStoryData, JiraStoryRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************

@dataclass
class AssigneeStoryRec:
    assignee: str = ''
    status: str = ''
    test_assignee: str = ''
    type: str = ''
    key: str = ''
    priority: str = ''
    points: int = 0
    summary: str = ''


class AssigneeDataRec:
    def __init__(self, assignee_in, jira_story_in):
        self.assignee: str = assignee_in
        self.stories: list[AssigneeStoryRec] = jira_story_in
        self.ws = None


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
        self.kcl_assignees: list[AssigneeDataRec] = []
        self.it_assignees: list[AssigneeDataRec] = []
        self.verisk_assignees: list[AssigneeDataRec] = []


# ******************************************************************************
# ******************************************************************************
# * Main
# ******************************************************************************
# ******************************************************************************
def main():

    print('\n\nStart Create Sprint Standup Assignee Spreadsheet')

    assignee_ss = AssigneeSS()

    jira_data = CsvJiraStoryData(Path.cwd() / 'Input files' / 'SGM - Jira - FAST Sprint Data (Jira).csv')
    # file_date = jira_data.ss_file_date
    if jira_data.stories:
        process_assignee_stories(jira_data.stories, assignee_ss.kcl_assignees, assignee_ss.it_assignees,
                                 assignee_ss.verisk_assignees)
        if assignee_ss.verisk_assignees or assignee_ss.kcl_assignees or assignee_ss.it_assignees:
            create_fast_standup_assignees_ss(assignee_ss, jira_data.ss_file_date)

    print('\nEnd Create Sprint Standup Assignee Spreadsheet')

    return None


# ==============================================================================
# ==============================================================================
# * Functions
# ==============================================================================
# ==============================================================================

# ==============================================================================
def process_assignee_stories(stories_in: list[JiraStoryRec], kcl_data: list[AssigneeDataRec],
                             it_data: list[AssigneeDataRec], verisk_data: list[AssigneeDataRec]) -> None:

    verisk_team = ('Madhava Krishna', 'Jacob Martinsen', 'Mark Davies', 'Rajiv Adepu', 'Aman Bhatt')

    it_team = ('Kris Dane', 'Ryan Akers', 'Tyler Herrada', 'Alex Kizub', 'Brian Collins', 'Dean Hill', 'Jay Huffman',
               'Kieran Ojakangas', 'Matanda Fatch', 'Sanita Gurung', 'Skylar Calvin', 'Evan Tindall', 'Goose Rodriguez',
               'Jim Richardson', 'Mallika Ramaswamy', 'David Scott', 'Paula Beruan', 'Terrence Lujin', 'Jason McQuinn',
               'Jim Vaughan')

    print('   Begin Processing Assignee Stories')
    for cur_story_rec in stories_in:
        assignee_story_rec = convert_jira_story_to_assignee_story(cur_story_rec)
        if assignee_story_rec.status != 'Done':
            match assignee_story_rec.status:
                case 'UAT':
                    if assignee_story_rec.test_assignee == 'Unassigned':
                        assignee = assignee_story_rec.assignee
                    else:
                        assignee = assignee_story_rec.test_assignee
                case 'QA':
                    if assignee_story_rec.test_assignee == 'Unassigned':
                        assignee = assignee_story_rec.assignee
                    else:
                        assignee = assignee_story_rec.test_assignee
                case _:
                    assignee = assignee_story_rec.assignee
            if assignee in verisk_team:
                update_assignee_data(verisk_data, assignee, assignee_story_rec)
            else:
                if assignee in it_team:
                    update_assignee_data(it_data, 'IT', assignee_story_rec)
                else:
                    update_assignee_data(kcl_data, assignee, assignee_story_rec)

    print('   Finished Processing ' + str(len(stories_in)) + ' Assignee Stories')

    return None


# ==============================================================================
def convert_jira_story_to_assignee_story(jira_story_rec: JiraStoryRec) -> AssigneeStoryRec:

    new_assignee_rec = AssigneeStoryRec()
    new_assignee_rec.assignee = jira_story_rec.assignee
    new_assignee_rec.status = jira_story_rec.status
    new_assignee_rec.test_assignee = jira_story_rec.test_assignee
    new_assignee_rec.type = jira_story_rec.type
    new_assignee_rec.key = jira_story_rec.key
    new_assignee_rec.priority = jira_story_rec.priority
    new_assignee_rec.points = jira_story_rec.points
    new_assignee_rec.summary = jira_story_rec.summary

    return new_assignee_rec


# ==============================================================================
def update_assignee_data(assignee_data: list[AssigneeDataRec], assignee_to_update: str,
                         assignee_story: AssigneeStoryRec) -> None:

    assignee_found = False
    if len(assignee_data) > 0:
        for cur_assignee_rec in assignee_data:
            if cur_assignee_rec.assignee == assignee_to_update:
                assignee_found = True
                cur_assignee_rec.stories.append(assignee_story)
                break
    if not assignee_found:
        new_assignee = AssigneeDataRec(assignee_to_update, [assignee_story])
        assignee_data.append(new_assignee)

    return None


# ==============================================================================
def update_it_data(it_data: list[AssigneeDataRec], assignee_story: AssigneeStoryRec) -> None:

    if len(it_data) > 0:
        it_data[0].stories.append(assignee_story)
    else:
        new_assignee = AssigneeDataRec('IT', [assignee_story])
        it_data.append(new_assignee)

    return None


# ==============================================================================
def create_fast_standup_assignees_ss(assignee_ss: AssigneeSS, file_date: str) -> None:
    print('\n   Creating FAST Standup Assignee spreadsheet')
    # create the spreadsheet workbook and formats for the spreadsheet
    assignee_ss.workbook = xlsxwriter.Workbook('Output files/' + file_date + ' Sprint Standup Assignees.xlsx')
    cell_formats = create_cell_formatting_options(assignee_ss.workbook)
    # Sort the assignees so that they are displayed in Alphabetic order
    assignee_ss.verisk_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee)
    assignee_ss.kcl_assignees.sort(key=lambda assignee_rec: assignee_rec.assignee)
    assignee_ss.it_assignees[0].stories.sort(key=lambda assignee_rec: assignee_rec.assignee)
    print('      Writing Verisk Assignees to Spreadsheet')
    for cur_assignee in assignee_ss.verisk_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.status, reverse=True)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)
    assignee_ws = assignee_ss.workbook.add_worksheet('Verisk Questions')
    print('      Writing IT Assignees to Spreadsheet')
    for cur_assignee in assignee_ss.it_assignees:
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee)
        create_assignee_ws_column_layout(assignee_ws, cell_formats)
        write_assignee_stories_to_ws(assignee_ws, cell_formats, cur_assignee.stories)
    print('      Writing Remaining Assignees to Spreadsheet')
    for cur_assignee in assignee_ss.kcl_assignees:
        cur_assignee.stories.sort(key=lambda assignee_rec: assignee_rec.status, reverse=True)
        assignee_ws = assignee_ss.workbook.add_worksheet(cur_assignee.assignee)
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
    cell_fmt.center_red_fmt = workbook.add_format({'align': 'center', 'font_color': 'green', 'bold': 1})
    cell_fmt.def_fmt = workbook.add_format({'align': 'left', 'indent': 1, 'text_wrap': 1})
    cell_fmt.table_label_fmt = workbook.add_format({'align': 'left', 'bold': 1, 'font_size': 14})

    return cell_fmt


# ==============================================================================
def create_assignee_ws_column_layout(worksheet, cell_fmts: Type[CellFormats]) -> None:
    # Set the column widths and default cell formatting for the Metrics tab
    # Setup Jira table layout
    worksheet.set_column('A:C', 25, cell_fmts.center_fmt)
    worksheet.set_column('D:G', 12, cell_fmts.center_fmt)
    worksheet.set_column('H:H', 80, cell_fmts.center_fmt)

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
        new_row_data = [cur_story.assignee,
                        cur_story.status,
                        cur_story.test_assignee,
                        cur_story.type,
                        cur_story.key,
                        cur_story.priority,
                        cur_story.points,
                        cur_story.summary]
        table_data.append(new_row_data)

    table_name = assignee_ws.name.replace(' ', '_')
    assignee_ws.add_table(assignee_story_tbl,
                           {'name': table_name,
                            'style': 'Table Style Medium 2',
                            'autofilter': True,
                            'first_column': False,
                            'data': table_data,
                            'columns': [{'header': 'Assignee', 'format': cell_fmts.left_fmt},
                                        {'header': 'Status', 'format': cell_fmts.left_fmt},
                                        {'header': 'Test Assignee', 'format': cell_fmts.left_fmt},
                                        {'header': 'Issue Type', 'format': cell_fmts.center_fmt},
                                        {'header': 'Issue Key', 'format': cell_fmts.center_fmt},
                                        {'header': 'Priority', 'format': cell_fmts.left_fmt},
                                        {'header': 'Story Points', 'format': cell_fmts.center_fmt},
                                        {'header': 'Summary', 'format': cell_fmts.left_fmt}]
                            })

    assignee_ws.conditional_format('C2:C40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="QA"',
                                    'format': cell_fmts.center_red_fmt})
    assignee_ws.conditional_format('C2:C40',
                                   {'type': 'formula',
                                    'criteria': '=$B2="UAT"',
                                    'format': cell_fmts.center_red_fmt})
    assignee_ws.conditional_format('B2:B40',
                                   {'type': 'text',
                                    'criteria': 'containsText',
                                    'value': 'QA',
                                    'format': cell_fmts.center_red_fmt})
    assignee_ws.conditional_format('B2:B40',
                                   {'type': 'text',
                                    'criteria': 'containsText',
                                    'value': 'UAT',
                                    'format': cell_fmts.center_red_fmt})
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
    main()
