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

# SGM Shared Module imports
sys.path.append('C:/Users/kap3309/OneDrive - Kansas City Life Insurance/PythonDev/Modules')

from kclGetFastTeamInfo import FASTTeamInfo
from kclGetFASTProjectInfo import FASTProjectInfo
from kclGetPIPlannedStoriesData_1 import PIPlannedStoryData
from kclGetJiraPlanningQueryData_1 import JiraPIStoryData
from kclGetCsvJiraStoryData import CsvJiraStoryData
from kclGetCsvJiraStoryDependencyData import CsvJiraStoryDependencyData


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
def main():
    print('\nStart Launcher')

    app_to_launch = get_app_to_launch()
    app_to_launch = 1
    match app_to_launch:
        case 0:
            project_info = FASTProjectInfo(Path.cwd() / 'Input files' / 'FAST Project Info.xlsx')
        case 1:
            pi_planned_story_data = PIPlannedStoryData(Path.cwd() / 'Input files' / 'PI_Plan_Q1_2023 - Planned.xlsx')
        case 2:
            jira_pi_story_data = JiraPIStoryData(Path.cwd() / 'Input files' / 'PI_Plan_Q1_2023 - Jira.xlsx')
        case 3:
            csv_file_story_data = CsvJiraStoryDependencyData(Path.cwd() / 'Input files' / 'Jira - Story Dependency Data.csv')
        case 4:
            csv_file_story_data = CsvJiraStoryData(Path.cwd() / 'Input files' / 'SGM - Jira - FAST Sprint Data (Jira).csv')
        case 5:
            fast_team_member_data = FASTTeamInfo(Path.cwd() / 'Input files' / 'FastTeamInfo.csv')
        case _:
            pass

    print('\nCompleted Test Modules')


if __name__ == "__main__":
    main()

# ******************************************************************************
# ******************************************************************************
# * Functions
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
