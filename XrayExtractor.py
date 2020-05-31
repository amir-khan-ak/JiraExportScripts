import json
import requests
import xlsxwriter
import sys


# url = 'http://191.2.5.28:8232/rest'
# user = "amir"
# pw = "!k#4n4miR"
# ContentType = {'Content-Type': 'application/json'}
# vsPath = 'C:\\XRAY\\Project-XRay-Test-Export.xlsx'

def writeToExcel(worksheet, row, column, data):
    worksheet.write(row, column, data)
    column += 1

    return worksheet, column


url = sys.argv[1] + '/rest'
user = sys.argv[2]
pw = sys.argv[3]
ContentType = {'Content-Type': 'application/json'}
vsPath = sys.argv[4]

# Login to Jira XRay
resource = 'auth/1/session'
payload = {"username": user, "password": pw}
resp = requests.post(url + '/' + resource,
                     data=json.dumps(payload),
                     headers=ContentType)

print('Login to Jira: ' + str(resp.status_code))

# Get test cases
JSESSIONID = resp.cookies._cookies['192.172.627.221']['/']['JSESSIONID'].value
ContentType = {'Content-Type': 'application/json', 'Cookie': 'JSESSIONID=' + JSESSIONID}
resource = 'raven/1.0/api/test?jql=project=XRAY'
resp = requests.get(url + '/' + resource,
                    headers=ContentType)
print('Getting XRay tests: ' + str(len(resp.json())))

################
# EXCEL
# Column Excel
ColumnsXLS = ['unique_id', 'type', 'name', 'step_type', 'step_description', 'test_type', 'product_areas',
              'covered_content',
              'designer', 'description', 'estimated_duration', 'owner', 'phase', 'user_tags', 'xray_id_udf']

workbook = xlsxwriter.Workbook(vsPath)

# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("manual tests")
column = 0
row = 0
unique_id = 0
_blank = ''

for col in ColumnsXLS:
    # write operation perform
    worksheet, column = writeToExcel(worksheet, row, column, col)


column = 0

for tests in resp.json():
    _name = 'api/2/issue/' + tests['key']
    _resp = requests.get(url + '/' + _name,
                         headers=ContentType)

    _testname = _resp.json()['fields']['summary']
    print('Test Name: ' + _testname)

    row = row + 1
    unique_id = unique_id + 1
    worksheet, column = writeToExcel(worksheet, row, column, unique_id)
    worksheet, column = writeToExcel(worksheet, row, column, 'test_manual')
    worksheet, column = writeToExcel(worksheet, row, column, _testname)
    worksheet, column = writeToExcel(worksheet, row, column, _blank)
    worksheet, column = writeToExcel(worksheet, row, column, _blank)
    worksheet, column = writeToExcel(worksheet, row, column, 'Acceptance')
    worksheet, column = writeToExcel(worksheet, row, column, _blank)
    worksheet, column = writeToExcel(worksheet, row, column, _blank)
    worksheet, column = writeToExcel(worksheet, row, column, 'khanamir@microfocus.com')
    if _resp.json()['fields']['description'] is not None:
        worksheet.write(row, column, _resp.json()['fields']['description'])
    column += 1
    worksheet, column = writeToExcel(worksheet, row, column, _blank)
    worksheet, column = writeToExcel(worksheet, row, column, 'khanamir@microfocus.com')
    worksheet, column = writeToExcel(worksheet, row, column, 'New')
    worksheet, column = writeToExcel(worksheet, row, column, 'Xray_Imported')
    worksheet, column = writeToExcel(worksheet, row, column, tests['key'])

    column = 0

    print('Total steps: ' + str(len(tests['definition']['steps'])))
    steps = tests['definition']['steps']
    for step in steps:
        if step['step'] is not None:
            column = 0
            row = row + 1
            unique_id = unique_id + 1
            worksheet, column = writeToExcel(worksheet, row, column, unique_id)
            worksheet, column = writeToExcel(worksheet, row, column, 'step')
            worksheet, column = writeToExcel(worksheet, row, column, _blank)
            worksheet, column = writeToExcel(worksheet, row, column, 'simple')
            worksheet, column = writeToExcel(worksheet, row, column, step['step']['raw'])

        if 'raw' in step['result']:
            if len(step['result']['raw']) > 1:
                column = 0
                row = row + 1
                unique_id = unique_id + 1
                worksheet, column = writeToExcel(worksheet, row, column, unique_id)
                worksheet, column = writeToExcel(worksheet, row, column, 'step')
                worksheet, column = writeToExcel(worksheet, row, column, _blank)
                worksheet, column = writeToExcel(worksheet, row, column, 'Validation')
                worksheet, column = writeToExcel(worksheet, row, column, step['result']['raw'])

    column = 0

# Logout to Jira XRay
resource = 'auth/1/session'
payload = {"username": user, "password": pw}
resp = requests.delete(url + '/' + resource,
                       headers=ContentType)

cookie = resp.cookies
print('Logout from Jira: ' + str(resp.status_code))

# CLOSE EXCEL
workbook.close()

print('EXPORT FILE SAVED UNDER: ' + vsPath)
