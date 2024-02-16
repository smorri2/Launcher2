#!/usr/bin/env python3


# ******************************************************************************
# ******************************************************************************
# * Imports
# ******************************************************************************
# ******************************************************************************

# Standard library imports
from dataclasses import dataclass, field
from pathlib import Path
from lxml import etree
from decimal import Decimal, getcontext


# Third party imports
import xlsxwriter


# local file imports


# SGM Shared Module imports


# ******************************************************************************
# ******************************************************************************
# * Class Declarations
# ******************************************************************************
# ******************************************************************************


@dataclass
class AccountEntry:
    policy_num: str = ''
    gl_entry_id: str = ''
    gl_entry_hdr_id: str = ''
    disbursement: str = ''
    account: str = ''
    amount: Decimal = Decimal('0.00')
    reversal: str = ''
    trans_type_desc: str = ''


@dataclass
class GLEntryHdrIDGroup:
    hdr_id: str = ''
    hdr_amount: Decimal = Decimal('0.00')
    entries: list[AccountEntry] = field(default_factory=list)


@dataclass
class VoucherFileInfo:
    cycle_date: str = ''
    file_count: int = 0
    total_credit_amount: Decimal = Decimal('0.00')
    total_debit_amount: Decimal = Decimal('0.00')
    number_of_records: int = 0


@dataclass
class InputData:
    file_info: VoucherFileInfo()
    files_to_process: list = field(default_factory=list)
    unmatched_accounting_entries: list = field(default_factory=list)


@dataclass
class OutputData:
    balanced_hdr_groups: list = field(default_factory=list)
    unbalanced_hdr_groups: list = field(default_factory=list)
    eft_transactions: list = field(default_factory=list)


@dataclass
class VoucherFileReviewSS:
    workbook = None
    unbalanced_ws = None
    balanced_ws = None
    detail_ws = None
    eft_ws = None
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

# ==============================================================================
# ==============================================================================
# === Functions
# ==============================================================================
# ==============================================================================


# ==============================================================================
# === Main
# ==============================================================================

def create_fast_voucher_review_spreadsheet() -> None:
    print('\n\nStart Create FAST Voucher Review Spreadsheet')

    input_data = get_input_data()
    if input_data.unmatched_accounting_entries:
        hdr_id_groups = process_unmatched_accounting_entries(input_data.unmatched_accounting_entries)
        output_data = process_header_groups(hdr_id_groups)
        output_data.eft_transactions = create_list_of_eft_transactions(input_data.unmatched_accounting_entries)
        create_voucher_file_review_spreadsheet(output_data, input_data)
    print('\nEnd Create FAST Voucher Review Spreadsheet')

    return None


# ==============================================================================
def get_input_data() -> InputData:

    file_info = VoucherFileInfo()
    input_data = InputData(file_info)

    input_data.files_to_process = get_files_to_process()
    if input_data.files_to_process:
        for cur_xml_file_path in input_data.files_to_process:
            root = read_and_parse_xml_in_file(cur_xml_file_path)
            if root is not None:
                get_accounting_entries_from_parsed_xml_data(root, input_data)

    return input_data


# ==============================================================================
def get_files_to_process() -> list[Path]:

    input_files: list[Path] = []
    input_files_path = Path(Path.cwd() / 'Input files' / 'Voucher files')
    for cur_file in Path(input_files_path).iterdir():
        input_files.append(cur_file)

    return input_files


# ==============================================================================
def read_and_parse_xml_in_file(xml_file_path: Path) -> etree.Element:

    print('\n  Parsing XML File data')
    root = None
    if xml_file_path.exists():
        with open(xml_file_path) as f:
            tree = etree.parse(f)
        root = tree.getroot()

    return root


# ==============================================================================
def get_accounting_entries_from_parsed_xml_data(root: etree.Element,
                                                input_data: InputData) -> None:

    num_acct_entries = 0
    if root is not None:
        print('\n  Getting Accounting Entries from XML')
        for cur_xtract_rpt in root:
            if cur_xtract_rpt.tag == 'GLExtractReport':
                new_acct_entry = process_cur_extract_rpt(cur_xtract_rpt)
                input_data.unmatched_accounting_entries.append(new_acct_entry)
                num_acct_entries += 1
            else:
                match cur_xtract_rpt.tag:
                    case 'CycleDate':
                        input_data.file_info.cycle_date = cur_xtract_rpt.text[0:10]
                    case 'FileCount':
                        input_data.file_info.file_count = int(cur_xtract_rpt.text)
                    case 'TotalCreditAmount':
                        input_data.file_info.total_credit_amount = Decimal(cur_xtract_rpt.text).quantize(Decimal('.01'))
                    case 'TotalDebitAmount':
                        input_data.file_info.total_debit_amount = Decimal(cur_xtract_rpt.text).quantize(Decimal('.01'))
                    case 'NumberOfRecords':
                        input_data.file_info.number_of_records = int(cur_xtract_rpt.text)
                    case _:
                        pass
    print(f'\n  Cycle Date => {input_data.file_info.cycle_date}')
    print(f'  File Count => {input_data.file_info.file_count}')
    print(f'  Number of Records => {input_data.file_info.number_of_records}')
    print(f'  Total Credit Amount => {input_data.file_info.total_credit_amount}')
    print(f'  Total Debit Amount => {input_data.file_info.total_debit_amount}')

    difference = input_data.file_info.total_credit_amount + input_data.file_info.total_debit_amount
    if difference != 0.0:
        print(f'  Credit/Debit Difference => {round(difference, 2)}')

    print(f'\n\n  Number of accounting entries found => {num_acct_entries}')
    print(f'')

    return None


# ==============================================================================
def process_cur_extract_rpt(cur_xtract_rpt: etree.Element) -> AccountEntry:
    acct_entry = AccountEntry()
    num_fields_found = 0
    for cur_elem in cur_xtract_rpt:
        match cur_elem.tag:
            case 'AccountNumber':
                acct_entry.account = cur_elem.text
                num_fields_found += 1
            case 'ConvertedAmount':
                acct_entry.amount = Decimal(cur_elem.text).quantize(Decimal('.01'))
                num_fields_found += 1
            case 'GLEntryID':
                acct_entry.gl_entry_id = cur_elem.text
                num_fields_found += 1
            case 'IsReversal':
                acct_entry.reversal = cur_elem.text
                num_fields_found += 1
            case 'GLEntryHdrID':
                acct_entry.gl_entry_hdr_id = cur_elem.text
                num_fields_found += 1
            case 'IsDisbursmentTxnRelated':
                acct_entry.disbursement = cur_elem.text
                num_fields_found += 1
            case 'PolicyNumber':
                acct_entry.policy_num = cur_elem.text
                num_fields_found += 1
            case 'TransactionTypeDescription':
                acct_entry.trans_type_desc = cur_elem.text
                num_fields_found += 1
            case _:
                pass
        if num_fields_found == 8:
            break

    return acct_entry


# ==============================================================================
def process_unmatched_accounting_entries(unmatched_accounting_entries: list[AccountEntry]) -> list[GLEntryHdrIDGroup]:

    hdr_ids: list[GLEntryHdrIDGroup] = []
    for cur_entry in unmatched_accounting_entries:
        add_cur_entry_to_matching_hdr_group(cur_entry, hdr_ids)

    return hdr_ids


# ==============================================================================
def add_cur_entry_to_matching_hdr_group(gl_entry: AccountEntry, hdr_ids: list[GLEntryHdrIDGroup]) -> None:
    found = False
    for cur_hdr_group in hdr_ids:
        if gl_entry.gl_entry_hdr_id == cur_hdr_group.hdr_id:
            found = True
            cur_hdr_group.hdr_amount = cur_hdr_group.hdr_amount + gl_entry.amount
            cur_hdr_group.entries.append(gl_entry)
            break
    if not found:
        new_hdr_group = GLEntryHdrIDGroup(gl_entry.gl_entry_hdr_id, gl_entry.amount, [gl_entry])
        # insert new_hdr_group at front of list because if the next entry to add has the same header id
        # it will match the group id at the front of the list.  Minimizing the amount of compares to find
        # a match.  If it does not match then it is likely a new header id and we will search the entire
        # list to be sure, before we add the new_hdr_group to the front of the list.
        hdr_ids.insert(0, new_hdr_group)

    return None


# ==============================================================================
def create_list_of_eft_transactions(unmatched_accounting_entries: list[AccountEntry]) -> list[AccountEntry]:
    eft_transactions_list = []

    for cur_entry in unmatched_accounting_entries:
        if cur_entry.account == '10020 - KCL_UnitMissEFT_9982':
            eft_transactions_list.append(cur_entry)

    return eft_transactions_list


# ==============================================================================
def process_header_groups(hdr_group_list: list[GLEntryHdrIDGroup]) -> OutputData:

    output_data = OutputData([], [], [])

    for cur_hdr_group in hdr_group_list:
        if cur_hdr_group.hdr_amount == 0.0:
            output_data.balanced_hdr_groups.append(cur_hdr_group)
        else:
            output_data.unbalanced_hdr_groups.append(cur_hdr_group)

    return output_data


# ===============================================================================
def create_voucher_file_review_spreadsheet(output_data: OutputData, input_data: InputData) -> None:
    if output_data is not None:
        voucher_ss = create_spreadsheet(input_data.file_info.cycle_date)
        write_unbalanced_header_groups_to_spreadsheet(voucher_ss, output_data.unbalanced_hdr_groups)
        write_balanced_header_groups_to_spreadsheet(voucher_ss, output_data.balanced_hdr_groups)
        write_eft_transactions_to_spreadsheet(voucher_ss, output_data.eft_transactions)
        write_account_entry_details_to_spreadsheet(voucher_ss, input_data.unmatched_accounting_entries)
        voucher_ss.workbook.close()

    return None


# ===============================================================================
def create_spreadsheet(cycle_date: str) -> VoucherFileReviewSS:

    # create the spreadsheet workbook and formats for the IPM Planning spreadsheet
    voucher_ss = create_ss_workbook_and_formats(cycle_date)

    # Set up the Unmatched Acct Entries worksheet tab to hold the Unmatched Accounting Entries
    voucher_ss.unbalanced_ws = voucher_ss.workbook.add_worksheet('Unbalanced Entries')
    voucher_ss.unbalanced_ws.set_column('A:A', 20)  # Policy Number
    voucher_ss.unbalanced_ws.set_column('B:B', 40)  # Account
    voucher_ss.unbalanced_ws.set_column('C:D', 16)  # Amount, Reversal
    voucher_ss.unbalanced_ws.set_column('E:E', 28)  # Transaction Type
    voucher_ss.unbalanced_ws.set_column('F:G', 55)  # GLEntryID, GLEntryHdrID

    # Set up the Matched Acct Entries worksheet tab to hold the Matched Accounting Entries
    voucher_ss.balanced_ws = voucher_ss.workbook.add_worksheet('Balanced Entries')
    voucher_ss.balanced_ws.set_column('A:A', 20)  # Policy Number
    voucher_ss.balanced_ws.set_column('B:B', 40)  # Account
    voucher_ss.balanced_ws.set_column('C:D', 16)  # Amount, Reversal
    voucher_ss.balanced_ws.set_column('E:E', 28)  # Transaction Type
    voucher_ss.balanced_ws.set_column('F:G', 55)  # GLEntryID, GLEntryHdrID

    # Set up the eft transacations worksheet tab to hold the detail of Accounting Entries
    voucher_ss.eft_ws = voucher_ss.workbook.add_worksheet('EFT Transactions')
    voucher_ss.eft_ws.set_column('A:B', 20)  # Policy Number, Entry Type
    voucher_ss.eft_ws.set_column('C:C', 40)  # Account
    voucher_ss.eft_ws.set_column('D:F', 18)  # Amount, Reversal, Disbursement
    voucher_ss.eft_ws.set_column('G:G', 28)  # Transaction Type
    voucher_ss.eft_ws.set_column('H:I', 55)  # GLEntryID, GLEntryHdrID

    # Set up the detail Acct Entries worksheet tab to hold the detail of Accounting Entries
    voucher_ss.detail_ws = voucher_ss.workbook.add_worksheet('Accounting Entry Detail')
    voucher_ss.detail_ws.set_column('A:B', 20)  # Policy Number, Entry Type
    voucher_ss.detail_ws.set_column('C:C', 40)  # Account
    voucher_ss.detail_ws.set_column('D:F', 18)  # Amount, Reversal, Disbursement
    voucher_ss.detail_ws.set_column('G:G', 28)  # Transaction Type
    voucher_ss.detail_ws.set_column('H:I', 55)  # GLEntryID, GLEntryHdrID

    return voucher_ss


# ==============================================================================
def create_ss_workbook_and_formats(cycle_date: str) -> VoucherFileReviewSS:
    # create the IPM Planning spreadsheet data structure and then create spreadsheet workbook
    voucher_ss = VoucherFileReviewSS()

    voucher_ss.workbook = xlsxwriter.Workbook('Output files/' + cycle_date + ' Voucher File Review.xlsx')

    font_size = 14
    # add predefined formats to be used for formatting cells in the spreadsheet
    voucher_ss.left_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 1
    })
    voucher_ss.left_bold_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    voucher_ss.left_lv2_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 4
    })
    voucher_ss.right_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'right',
        'indent': 1
    })
    voucher_ss.center_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
    })
    voucher_ss.center_red_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF0000'
    })
    voucher_ss.center_orange_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#FF8000'
    })
    voucher_ss.center_green_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
        'font_color': '#00CC66'
    })
    voucher_ss.percent_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'right',
        'indent': 6,
        'num_format': '0%'
    })
    voucher_ss.header_fmt = voucher_ss.workbook.add_format({
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
    voucher_ss.last_row_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bottom': 6
    })
    voucher_ss.last_row_name_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'left',
        'indent': 1,
        'bottom': 6
    })
    voucher_ss.totals_fmt = voucher_ss.workbook.add_format({
        'font_size': font_size,
        'align': 'center',
        'bold': 1,
    })

    return voucher_ss


# ==============================================================================
def write_unbalanced_header_groups_to_spreadsheet(voucher_ss: VoucherFileReviewSS,
                                                  unbalanced_hdr_groups: list[GLEntryHdrIDGroup]) -> None:
    header_row = 1

    # Write the Assignee Totals Header
    voucher_ss.unbalanced_ws.write(f'A{header_row}', 'Policy Number', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'B{header_row}', 'Account', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'C{header_row}', 'Amount', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'D{header_row}', 'Reversal', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'E{header_row}', 'Transaction Type', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'F{header_row}', 'GLEntryID', voucher_ss.header_fmt)
    voucher_ss.unbalanced_ws.write(f'G{header_row}', 'GLEntryHdrID', voucher_ss.header_fmt)

    # Write the Unmatched Accounting Entries to Unmatched Entries worksheet
    ws_row = 1
    for cur_hdr_group in unbalanced_hdr_groups:
        for cur_entry in cur_hdr_group.entries:
            voucher_ss.unbalanced_ws.write(ws_row, 0, cur_entry.policy_num, voucher_ss.left_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 1, cur_entry.account, voucher_ss.left_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 2, cur_entry.amount, voucher_ss.right_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 3, cur_entry.reversal, voucher_ss.center_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 4, cur_entry.trans_type_desc, voucher_ss.left_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 5, cur_entry.gl_entry_id, voucher_ss.left_fmt)
            voucher_ss.unbalanced_ws.write(ws_row, 6, cur_entry.gl_entry_hdr_id, voucher_ss.left_fmt)
            ws_row += 1
        ws_row += 1

    return None


# ==============================================================================
def write_balanced_header_groups_to_spreadsheet(voucher_ss: VoucherFileReviewSS,
                                                balanced_hdr_groups: list[GLEntryHdrIDGroup]) -> None:
    header_row = 1

    # Write the Assignee Totals Header
    voucher_ss.balanced_ws.write(f'A{header_row}', 'Policy Number', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'B{header_row}', 'Account', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'C{header_row}', 'Amount', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'D{header_row}', 'Reversal', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'E{header_row}', 'Transaction Type', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'F{header_row}', 'GLEntryID', voucher_ss.header_fmt)
    voucher_ss.balanced_ws.write(f'G{header_row}', 'GLEntryHdrID', voucher_ss.header_fmt)

    # Write the Unmatched Accounting Entries to Unmatched Entries worksheet
    ws_row = 1
    for cur_hdr_group in balanced_hdr_groups:
        for cur_entry in cur_hdr_group.entries:
            voucher_ss.balanced_ws.write(ws_row, 0, cur_entry.policy_num, voucher_ss.left_fmt)
            voucher_ss.balanced_ws.write(ws_row, 1, cur_entry.account, voucher_ss.left_fmt)
            voucher_ss.balanced_ws.write(ws_row, 2, cur_entry.amount, voucher_ss.right_fmt)
            voucher_ss.balanced_ws.write(ws_row, 3, cur_entry.reversal, voucher_ss.center_fmt)
            voucher_ss.balanced_ws.write(ws_row, 4, cur_entry.trans_type_desc, voucher_ss.left_fmt)
            voucher_ss.balanced_ws.write(ws_row, 5, cur_entry.gl_entry_id, voucher_ss.left_fmt)
            voucher_ss.balanced_ws.write(ws_row, 6, cur_entry.gl_entry_hdr_id, voucher_ss.left_fmt)
            ws_row += 1
        ws_row += 1
    return None


# ==============================================================================
def write_eft_transactions_to_spreadsheet(voucher_ss: VoucherFileReviewSS, eft_transactions_detail: list[AccountEntry]) -> None:
    header_row = 1

    # Write the Assignee Totals Header
    voucher_ss.eft_ws.write(f'A{header_row}', 'Policy Number', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'B{header_row}', 'Entry Type', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'C{header_row}', 'Account', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'D{header_row}', 'Amount', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'E{header_row}', 'Reversal', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'F{header_row}', 'Disbursement Txn Related', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'G{header_row}', 'Transaction Type', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'H{header_row}', 'GLEntryID', voucher_ss.header_fmt)
    voucher_ss.eft_ws.write(f'I{header_row}', 'GLEntryHdrID', voucher_ss.header_fmt)

    # Write the Unmatched Accounting Entries to Unmatched Entries worksheet
    ws_row = 1
    for cur_entry in eft_transactions_detail:
        voucher_ss.eft_ws.write(ws_row, 0, cur_entry.policy_num, voucher_ss.left_fmt)
        if cur_entry.amount > 0.0:
            voucher_ss.eft_ws.write(ws_row, 1, 'Debit', voucher_ss.left_fmt)
        else:
            voucher_ss.eft_ws.write(ws_row, 1, 'Credit', voucher_ss.left_fmt)
        voucher_ss.eft_ws.write(ws_row, 2, cur_entry.account, voucher_ss.left_fmt)
        voucher_ss.eft_ws.write(ws_row, 3, cur_entry.amount, voucher_ss.right_fmt)
        voucher_ss.eft_ws.write(ws_row, 4, cur_entry.reversal, voucher_ss.center_fmt)
        voucher_ss.eft_ws.write(ws_row, 5, cur_entry.disbursement, voucher_ss.center_fmt)
        voucher_ss.eft_ws.write(ws_row, 6, cur_entry.trans_type_desc, voucher_ss.left_fmt)
        voucher_ss.eft_ws.write(ws_row, 7, cur_entry.gl_entry_id, voucher_ss.left_fmt)
        voucher_ss.eft_ws.write(ws_row, 8, cur_entry.gl_entry_hdr_id, voucher_ss.left_fmt)
        ws_row += 1

    return None


# ==============================================================================
def write_account_entry_details_to_spreadsheet(voucher_ss: VoucherFileReviewSS,
                                               acct_entry_detail: list[AccountEntry]) -> None:
    header_row = 1

    # Write the Assignee Totals Header
    voucher_ss.detail_ws.write(f'A{header_row}', 'Policy Number', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'B{header_row}', 'Entry Type', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'C{header_row}', 'Account', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'D{header_row}', 'Amount', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'E{header_row}', 'Reversal', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'F{header_row}', 'Disbursement Txn Related', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'G{header_row}', 'Transaction Type', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'H{header_row}', 'GLEntryID', voucher_ss.header_fmt)
    voucher_ss.detail_ws.write(f'I{header_row}', 'GLEntryHdrID', voucher_ss.header_fmt)

    # Write the Unmatched Accounting Entries to Unmatched Entries worksheet
    ws_row = 1
    for cur_entry in acct_entry_detail:
        voucher_ss.detail_ws.write(ws_row, 0, cur_entry.policy_num, voucher_ss.left_fmt)
        if cur_entry.amount > 0.0:
            voucher_ss.detail_ws.write(ws_row, 1, 'Debit', voucher_ss.left_fmt)
        else:
            voucher_ss.detail_ws.write(ws_row, 1, 'Credit', voucher_ss.left_fmt)
        voucher_ss.detail_ws.write(ws_row, 2, cur_entry.account, voucher_ss.left_fmt)
        voucher_ss.detail_ws.write(ws_row, 3, cur_entry.amount, voucher_ss.right_fmt)
        voucher_ss.detail_ws.write(ws_row, 4, cur_entry.reversal, voucher_ss.center_fmt)
        voucher_ss.detail_ws.write(ws_row, 5, cur_entry.disbursement, voucher_ss.center_fmt)
        voucher_ss.detail_ws.write(ws_row, 6, cur_entry.trans_type_desc, voucher_ss.left_fmt)
        voucher_ss.detail_ws.write(ws_row, 7, cur_entry.gl_entry_id, voucher_ss.left_fmt)
        voucher_ss.detail_ws.write(ws_row, 8, cur_entry.gl_entry_hdr_id, voucher_ss.left_fmt)
        ws_row += 1

    return None


if __name__ == "__main__":
    create_fast_voucher_review_spreadsheet()
