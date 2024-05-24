import gspread
from google.oauth2.service_account import Credentials

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]

creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "14m8OvTbAHmh3tUYIqWXgUFxOZeg0iFkb7PAQRwahxGM"

workbook = client.open_by_key(sheet_id)

# values_list = sheet.sheet1.row_values(1)

# print(values_list)

# list the sheets in a workbook
# sheets = workbook.worksheets()

# sheets = map(lambda x : x.title,workbook.worksheets())

# print(list(sheets))

# sheet = workbook.worksheet('Hello world!')

# sheet.update_title('Sheet1')

sheet = workbook.worksheet('Sheet1')

# update the values in the cell using acell
# sheet.update_acell('A1', "column1")

# update the values in the cell using cell
# sheet.update_cell(1, 1, "column1")

# find the value in cell
# value = sheet.acell('A1').value
# value = sheet.cell(1, 2).value

# print(value)

# finding the cell based on value in the cell
# cell = sheet.find('column1')

# print(cell.row, cell.col)

# Basic formatting for cells
sheet.format('A1', {"textFormat":{"bold":True}})