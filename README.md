# Export data from a Jira on-premise installation to Excel file
# JiraExportScripts
This repository contains scripts to extract data from Jira on-premise into Excel file in preperation to upload to ALM Octane. 

# ----------------------------------------------------------------------------------------------
# XRayExtractor.py
This python script connects through REST API to a Jira on-premise instlalation and extract test case information to a pre-formatted Excel file in preperation to upload tests to ALM Octane. The script is intented to be used as an one time activity and not as a synchronization solution. In order to use an synchronization solution, check out the following link: https://marketplace.microfocus.com/appdelivery/content/micro-focus-connect-core

# Pre-requisites
- Download & Install Python 3.8 or higher: https://www.python.org/downloads/
- Install xlsxwriter library - https://pypi.org/project/XlsxWriter/
- Install requests library - https://pypi.org/project/requests/2.7.0/

# Usage
Using commandline run:
cd "directory where importUsers.py is located"
python XRayExtractor.py <jira_url> <jira_user> <jira_user_password> <path_to_save_excel_file.xls>
  
# Parameters
- <jira_url> Jira on-premise instance URL in the exact format as follow: 'http(s)://serverhost:port'. Don't add the '/' at the end of the URL.
- <jira_user> Jira user with permissions to read test cases, issues and additional fields.
- <jira_user_password> Password for the jira user in order to authenticate through the REST API.
- <path_to_save_excel_file.xls> the full path including the file extension where the excel file should be saved.

# Example: 
cd C:\XRayExtractor
C:\XRayExtractedData>python XRayExtractor.py "https://112.133.5.18:8481" "admin" "lke98dje3q+!Dkxv3" "C:\\XRayExtractedData\\XRayExtractedTests.xlsx"

# ----------------------------------------------------------------------------------------------
