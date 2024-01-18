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
from tqdm import tqdm
import xlsxwriter


# SGM Shared Module imports
from kclGetFASTCaseStatusData_1 import FastCaseStatusData
from kclGetFASTPolicyStatusData_1 import FASTPolicyStatusData, PolicyStatusRec
from kclGetProdSummaryAddlFieldsData_1 import ProdSummaryAddlFieldsData, AddlFieldsRec


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************

class AddlFieldsSS:
    def __init__(self):
        self.workbook = None
        self.data_ws = None
        self.header_fmt = None
        self.left_fmt = None
        self.right_fmt = None
        self.percent_fmt = None
        self.acct_fmt = None
        self.center_fmt = None
        self.addl_fields_tbl: str = ''


class PolicyStatusSS:
    def __init__(self):
        self.workbook = None
        self.data_ws = None
        self.header_fmt = None
        self.left_fmt = None
        self.right_fmt = None
        self.percent_fmt = None
        self.date_fmt = None
        self.center_fmt = None
        self.policy_status_tbl: str = ''


@dataclass()
class InputData:
    success: bool = False
    case_status_data: FastCaseStatusData = None
    policy_status_data: FASTPolicyStatusData = None
    addl_fields_ws_data: ProdSummaryAddlFieldsData = None

# ******************************************************************************
# ******************************************************************************
# # * Functions
# ******************************************************************************
# ******************************************************************************


# ==============================================================================
# ==============================================================================
# = Main
# ==============================================================================
# ==============================================================================
def case_status_true_up():

    print('\nStart Production Summary Case Status True Up')

    input_data = get_input_data()

    if input_data.success:
        compare_case_status_data_to_addl_fields_ws_status_data(input_data)
        write_addl_fields_ws_status_data(input_data.addl_fields_ws_data)
        write_policy_status_data_policy_ws_status_data(input_data.policy_status_data)

    print('\nCompleted Production Summary Case Status True Up')

    return None


# ==============================================================================
def get_input_data() -> InputData:
    input_data = InputData()
    input_data.case_status_data = get_fast_case_status_data_to_process()
    if input_data.case_status_data is not None:
        input_data.policy_status_data = get_fast_policy_status_data_to_process()
        if input_data.policy_status_data is not None:
            input_data.addl_fields_ws_data = get_prod_summary_addl_fields_ws_data_to_review()
            if input_data.addl_fields_ws_data is not None:
                input_data.success = True

    return input_data


# ==============================================================================
def get_fast_case_status_data_to_process() -> FastCaseStatusData:
    case_status_data = None
    # get the FAST Case status data to process
    # build the path to the Input folder where the CaseHdrObject.csv file resides
    # CaseHdrObject.csv contains the FAST Case status data for the FAST Cases to process
    fast_case_status_file_path = Path.cwd() / 'Input files' / 'CaseHdrObject.csv'
    if fast_case_status_file_path.exists():
        case_status_data = FastCaseStatusData(fast_case_status_file_path)
        if case_status_data is None:
            print('****** Error Getting Case Status Data from CaseHdrObject.csv ******')
    else:
        print('****** Error CaseHdrObject.csv not found ******')

    return case_status_data


# ==============================================================================
def get_fast_policy_status_data_to_process() -> FASTPolicyStatusData:

    # get the FAST Policy status data to process
    # build the path to the Input folder where the PolicyHdrObject.csv file resides
    # CaseHdrObject.csv contains the FAST Case status data for the FAST Cases to process
    fast_policy_status_file_path = Path.cwd() / 'Input files' / 'PolicyHdrObject.csv'
    policy_status_data = FASTPolicyStatusData(fast_policy_status_file_path)

    if policy_status_data is None:
        print('****** Error Getting Policy Status Data from PolicyHdrObject.csv ******')

    return policy_status_data


# ==============================================================================
def get_prod_summary_addl_fields_ws_data_to_review() -> ProdSummaryAddlFieldsData:

    # get the FAST Case status data to process
    # build the path to the Input folder where the CaseHdrObject.csv file resides
    # CaseHdrObject.csv contains the FAST Case status data for the FAST Cases to process
    fast_prod_summary_file_path = Path.cwd() / 'Input files' / 'FAST Production Summary.xlsx'
    prod_sum_addl_items_data = ProdSummaryAddlFieldsData(fast_prod_summary_file_path)

    return prod_sum_addl_items_data


# ==============================================================================
def compare_case_status_data_to_addl_fields_ws_status_data(input_data: InputData) -> None:

    # Process through the case_status_data list of data and compare to addl_fields status data
    # progress via the tqdm Progress Bar
    print('   Start Comparing Addl Fields Case Status to Case Header Case Status')
    pbar = tqdm(total=len(input_data.case_status_data.case_list), desc='      Case Status Compare ',
                ncols=120, bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt}")

    for cur_case_status_rec in input_data.case_status_data:
        addl_fields_case = input_data.addl_fields_ws_data.search_for_case_num(cur_case_status_rec.case_number)
        if addl_fields_case:
            addl_fields_case.updated_case_status = cur_case_status_rec.status  # remove after testing complete
            addl_fields_case.case_status = cur_case_status_rec.status
        pbar.update()  # Update the Progress Bar
    pbar.close()  # close the Progress Bar display
    print('   Finished Comparing Addl Fields Case Status to Case Header Case Status')

    return None


# ==============================================================================
def create_addl_fields_wb_and_formats() -> AddlFieldsSS:
    # create the sprint report spreadsheet data structure and then create spreadsheet workbook
    addl_fields_ss = AddlFieldsSS()

    addl_fields_ss.workbook = xlsxwriter.Workbook('Output files/Addl Fields Compare.xlsx')

    # add predefined formats to be used for formatting cells in the spreadsheet
    addl_fields_ss.left_fmt = addl_fields_ss.workbook.add_format({
        'font_size': 11,
        'align': 'left',
        'indent': 1
    })
    addl_fields_ss.left_bold_fmt = addl_fields_ss.workbook.add_format({
        'font_size': 11,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    addl_fields_ss.right_fmt = addl_fields_ss.workbook.add_format({
        'font_size': 11,
        'align': 'right',
        'indent': 1
    })
    addl_fields_ss.center_fmt = addl_fields_ss.workbook.add_format({
        'font_size': 11,
        'align': 'center',
    })
    addl_fields_ss.acct_fmt = addl_fields_ss.workbook.add_format({
        'font_size': 11,
        'align': 'right',
        'indent': 1,
        'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    })
    addl_fields_ss.header_fmt = addl_fields_ss.workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 12,
        'font_color': 'white',
        'align': 'center',
        'bold': 1,
        'bg_color': '#4472C4',
        'pattern': 1,
        'border': 1
    })

    return addl_fields_ss


# ==============================================================================
def calc_table_starting_and_ending_cells(top_row: int, left_col, right_col, num_data_rows) -> str:
    top_left_cell = left_col + str(top_row)
    bot_right_cell = right_col + str(top_row + num_data_rows + 1)
    table_coordinates = top_left_cell + ':' + bot_right_cell

    return table_coordinates


# ==============================================================================
def create_addl_fields_tab_worksheet(add_fields_ss: AddlFieldsSS, cases: list[AddlFieldsRec]) -> None:

    print('      ** Writing Addl Fields spreadsheet tab')
    add_fields_ss.data_ws = add_fields_ss.workbook.add_worksheet('Addl Fields')

    # Setup Details table layout
    add_fields_ss.data_ws.set_column('A:A', 14, add_fields_ss.left_fmt)
    add_fields_ss.data_ws.set_column('B:B', 40, add_fields_ss.center_fmt)
    add_fields_ss.data_ws.set_column('C:D', 14, add_fields_ss.center_fmt)
    add_fields_ss.data_ws.set_column('E:E', 30, add_fields_ss.center_fmt)
    add_fields_ss.data_ws.set_column('F:G', 12, add_fields_ss.center_fmt)

    # ******************************************************************
    # Set Addl Fields Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    add_fields_ss.addl_fields_tbl = calc_table_starting_and_ending_cells(1, 'A', 'G', len(cases) - 1)

    return None


# ==============================================================================
def write_addl_fields_data_to_ss(add_fields_ss: AddlFieldsSS, case_data: list[AddlFieldsRec]) -> None:

    table_data = []
    for cur_case in case_data:
        new_row_data = [cur_case.case_number,
                        cur_case.case_status,
                        # cur_case.updated_case_status,
                        cur_case.modal_prem,
                        cur_case.agent,
                        cur_case.agent_name,
                        cur_case.agency,
                        cur_case.company]
        table_data.append(new_row_data)

    add_fields_ss.data_ws.add_table(add_fields_ss.addl_fields_tbl,
                                    {'name': 'addl_fields_table',
                                     'style': 'Table Style Light 15',
                                     'autofilter': True,
                                     'first_column': False,
                                     'banded_rows': False,
                                     'data': table_data,
                                     'columns': [{'header': 'Policy', 'format': add_fields_ss.left_fmt},
                                                 {'header': 'Case Status', 'format': add_fields_ss.left_fmt},
                                                 {'header': 'Modal Prem', 'format': add_fields_ss.acct_fmt},
                                                 {'header': 'Agent', 'format': add_fields_ss.right_fmt},
                                                 {'header': 'Agent Name', 'format': add_fields_ss.left_fmt},
                                                 {'header': 'Agency', 'format': add_fields_ss.center_fmt},
                                                 {'header': 'Grange/KCL', 'format': add_fields_ss.left_fmt}]
                                     })

    return None


# ==============================================================================
def write_addl_fields_ws_status_data(addl_fields_ws_data: ProdSummaryAddlFieldsData) -> None:
    print('\n   Creating Addl Fields Compare spreadsheet')
    # create the spreadsheet workbook and formats for the addl fields spreadsheet
    addl_fields_ss = create_addl_fields_wb_and_formats()

    # create the worksheet and table layouts for the Addl Fields Data tab in the Addl Fields Compare spreadsheet
    create_addl_fields_tab_worksheet(addl_fields_ss, addl_fields_ws_data.cases)

    # write the Addl Fields Data to Addl Fields tab of spreadsheet
    write_addl_fields_data_to_ss(addl_fields_ss, addl_fields_ws_data.cases)

    addl_fields_ss.workbook.close()
    print('   Completed Addl Fields Compare Spreadsheet')

    return None


# ==============================================================================
def create_policy_status_wb_and_formats() -> PolicyStatusSS:
    # create the sprint report spreadsheet data structure and then create spreadsheet workbook
    policy_status_ss = PolicyStatusSS()

    policy_status_ss.workbook = xlsxwriter.Workbook('Output files/Policy Status.xlsx')

    # add predefined formats to be used for formatting cells in the spreadsheet
    policy_status_ss.left_fmt = policy_status_ss.workbook.add_format({
        'font_size': 11,
        'align': 'left',
        'indent': 1
    })
    policy_status_ss.left_bold_fmt = policy_status_ss.workbook.add_format({
        'font_size': 11,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    policy_status_ss.right_fmt = policy_status_ss.workbook.add_format({
        'font_size': 11,
        'align': 'right',
        'indent': 1
    })
    policy_status_ss.center_fmt = policy_status_ss.workbook.add_format({
        'font_size': 11,
        'align': 'center',
    })
    policy_status_ss.date_fmt = policy_status_ss.workbook.add_format({
        'font_size': 11,
        'align': 'center',
        'num_format': 'mm/dd/yy;@'
    })
    policy_status_ss.header_fmt = policy_status_ss.workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 12,
        'font_color': 'white',
        'align': 'center',
        'bold': 1,
        'bg_color': '#4472C4',
        'pattern': 1,
        'border': 1
    })

    return policy_status_ss


# ==============================================================================
def create_policy_status_tab_worksheet(policy_status_ss: PolicyStatusSS, policies: list[PolicyStatusRec]) -> None:

    print('      ** Writing Policy Status spreadsheet tab')
    policy_status_ss.data_ws = policy_status_ss.workbook.add_worksheet('Policy Status')

    # Setup Details table layout
    policy_status_ss.data_ws.set_column('A:A', 16, policy_status_ss.left_fmt)
    policy_status_ss.data_ws.set_column('B:H', 20, policy_status_ss.left_fmt)
    # policy_status_ss.data_ws.set_column('C:C', 12, policy_status_ss.center_fmt)
    # policy_status_ss.data_ws.set_column('D:H', 20, policy_status_ss.center_fmt)

    # ******************************************************************
    # Set Addl Fields Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    policy_status_ss.policy_status_tbl = calc_table_starting_and_ending_cells(1, 'A', 'H', len(policies) - 1)

    return None


# ==============================================================================
def write_policy_status_data_to_ss(policy_status_ss: PolicyStatusSS, policy_data: list[PolicyStatusRec]) -> None:

    policy_data.sort(key=lambda policy_rec: policy_rec.timestamp, reverse=True)

    table_data = []
    for cur_policy in policy_data:
        new_row_data = [cur_policy.policy_number,
                        cur_policy.status,
                        cur_policy.timestamp,
                        cur_policy.issue_state,
                        cur_policy.application_date,
                        cur_policy.app_received_date,
                        cur_policy.policy_effective_date,
                        cur_policy.app_type]
        table_data.append(new_row_data)

    policy_status_ss.data_ws.add_table(policy_status_ss.policy_status_tbl,
                                       {'name': 'policy_status_table',
                                        'style': 'Table Style Light 15',
                                        'autofilter': True,
                                        'first_column': False,
                                        'banded_rows': False,
                                        'data': table_data,
                                        'columns': [{'header': 'Policy Number', 'format': policy_status_ss.left_fmt},
                                                    {'header': 'Policy Status', 'format': policy_status_ss.left_fmt},
                                                    {'header': 'Timestamp', 'format': policy_status_ss.date_fmt},
                                                    {'header': 'Issue State', 'format': policy_status_ss.center_fmt},
                                                    {'header': 'App Received Date',
                                                     'format': policy_status_ss.date_fmt},
                                                    {'header': 'Application Date', 'format': policy_status_ss.date_fmt},
                                                    {'header': 'Policy Effective Date',
                                                     'format': policy_status_ss.date_fmt},
                                                    {'header': 'App Type', 'format': policy_status_ss.left_fmt}]
                                        })

    return None


# ==============================================================================
def write_policy_status_data_policy_ws_status_data(policy_status_data: FASTPolicyStatusData) -> None:
    print('\n   Creating Policy Status spreadsheet')
    # create the spreadsheet workbook and formats for the policy status spreadsheet
    policy_status_ss = create_policy_status_wb_and_formats()

    # create the worksheet and table layouts for the Addl Fields Data tab in the Addl Fields Compare spreadsheet
    create_policy_status_tab_worksheet(policy_status_ss, policy_status_data.policies)

    # write the Addl Fields Data to Addl Fields tab of spreadsheet
    write_policy_status_data_to_ss(policy_status_ss, policy_status_data.policies)

    policy_status_ss.workbook.close()
    print('   Completed Policy Status Spreadsheet')

    return None


if __name__ == "__main__":
    case_status_true_up()
