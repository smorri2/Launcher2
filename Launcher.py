#!/usr/bin/env python3


# **********************************************************************************************************************
# **********************************************************************************************************************
# * Imports
# **********************************************************************************************************************
# **********************************************************************************************************************

# Standard library imports
import sys
sys.path.append('C:/Users/kap3309/OneDrive - Kansas City Life Insurance/PythonDev/Modules')
# Third party imports

# local application imports
from Create_FAST_IPM_Planning_Report import create_fast_ipm_planning_spreadsheet
from Create_FAST_Standup_Assignees_Spreadsheet import create_standup_assignees_spreadsheet
from Create_FAST_Sprint_Report import create_sprint_report
from FastVoucherFileReview import create_fast_voucher_review_spreadsheet
from Create_PI_Metrics import create_pi_planning_metrics
from Create_FAST_CS_Letter_Report import create_cs_letter_report
from Create_FAST_Control_Report_Tracking import create_fast_control_report_tracking
from Plan_Program_Increment import plan_program_increment
from Sprint_Story_Dependencies import sprint_story_dependencies


# SGM Shared Module imports

# ******************************************************************************
# ******************************************************************************
# = functions
# ******************************************************************************
# ******************************************************************************


# ******************************************************************************
# ******************************************************************************
# = Main
# ******************************************************************************
# ******************************************************************************

# ==============================================================================
def get_app_to_launch() -> int:
    app_to_launch: int = 0
    valid_input: bool = False

    while not valid_input:
        print('\n')
        print('   ******************************************************************')
        print('   **                                                             ***')
        print('   **    Select the App to Launch                                 ***')
        print('   **                                                             ***')
        print('   **         1 - Create FAST Standup Assignees Spreadsheet       ***')
        print('   **         2 - Create FAST Sprint Report                       ***')
        print('   **         3 - Create FAST Voucher Review Spreadsheet          ***')
        print('   **         4 - Create FAST IPM Spreadsheet                     ***')
        print('   **         5 - Create PI Planning Metrics Spreadsheet          ***')
        print('   **         6 - Create CS Letter Report Spreadsheet             ***')
        print('   **         7 - Create Control Report Tracking Spreadsheet      ***')
        print('   **         8 - Create Plan Program Increment Spreadsheet       ***')
        print('   **         9 - Create Sprint Story Dependencies Spreadsheet    ***')
        print('   **         0 - Quit                                            ***')
        print('   **                                                             ***')
        print('   ******************************************************************')

        user_input = input('\n   Enter Number to Launch  ==> ')
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
            case '4':
                app_to_launch = int(user_input)
                valid_input = True
            case '5':
                app_to_launch = int(user_input)
                valid_input = True
            case '6':
                app_to_launch = int(user_input)
                valid_input = True
            case '7':
                app_to_launch = int(user_input)
                valid_input = True
            case '8':
                app_to_launch = int(user_input)
                valid_input = True
            case '9':
                app_to_launch = int(user_input)
                valid_input = True
            case '0':
                app_to_launch = int(user_input)
                valid_input = True
            case _:
                valid_input = False
                print('\n\n\n   Invalid App Number, valid App Numbers are between 1 & 4 inclusive')

    return app_to_launch


# ==============================================================================
def main():
    print('\nStart Launcher')

    done = False
    while not done:
        app_to_launch = get_app_to_launch()
        match app_to_launch:
            case 0:
                done = True
            case 1:
                create_standup_assignees_spreadsheet()
            case 2:
                create_sprint_report()
            case 3:
                create_fast_voucher_review_spreadsheet()
            case 4:
                create_fast_ipm_planning_spreadsheet()
            case 5:
                create_pi_planning_metrics()
            case 6:
                create_cs_letter_report()
            case 7:
                create_fast_control_report_tracking()
            case 8:
                plan_program_increment()
            case 9:
                sprint_story_dependencies()
            case _:
                pass

    print('\nCompleted Launcher')


if __name__ == "__main__":
    main()

# ******************************************************************************
# ******************************************************************************
# * Functions
# ******************************************************************************
# ******************************************************************************
