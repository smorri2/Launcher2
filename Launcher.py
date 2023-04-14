#!/usr/bin/env python3


# **********************************************************************************************************************
# **********************************************************************************************************************
# * Imports
# **********************************************************************************************************************
# **********************************************************************************************************************

# Standard library imports
from pathlib import Path

# Third party imports

# local application imports
from Create_FAST_Standup_Assignees_Spreadsheet import create_standup_assignees_spreadsheet

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
        print('   **         3 - Create PI Planning Metrics                      ***')
        print('   **         4 - Create FAST IPM Spreadsheet                     ***')
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
            case _:
                valid_input = False
                print('\n\n\n   Invalid PI Number, valid App Numbers are between 1 & 4 inclusive')

    return app_to_launch


# ==============================================================================
def main():
    print('\nStart Launcher')

    app_to_launch = get_app_to_launch()
    app_to_launch = 1
    match app_to_launch:
        case 1:
            create_standup_assignees_spreadsheet()
        case 2:
            pass
        case 3:
            pass
        case 4:
            pass
        case 5:
            pass
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
