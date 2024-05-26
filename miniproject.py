# mini project for googlesheetsAPI tutorial in python
import gspread
from google.oauth2.service_account import Credentials

scopes = [
    'https://www.googleapis.com/auth/spreadsheets'
]

creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "14m8OvTbAHmh3tUYIqWXgUFxOZeg0iFkb7PAQRwahxGM"

workbook = client.open_by_key(sheet_id)

values = [
    ["Name", "Price", "Quantity"],
    ["Basketball", 29.99, 1],
    ["Jeans", 39.99, 4],
    ["Soap", 7.99, 3]
]

worksheet_list = map(lambda x: x.title, workbook.worksheets())

new_worksheet_name = "Values"

if new_worksheet_name in worksheet_list:
    sheet = workbook.worksheet(new_worksheet_name)
else:
    sheet = workbook.add_worksheet(new_worksheet_name, rows=10, cols=10)
    
sheet.clear()   # Clear all the contents in the sheet

# inefficient way to update one by one cell in for loop
# for i, row in enumerate(values):
#     for j, value in enumerate(row):
#         sheet.update_cell(i+1, j+1, value)  # sheet index always starts with 1, 1

# Efficient way to update a range of values
sheet.update(f'A1:C{len(values)}', values)

# sheet formatting titles
sheet.format("A1:C1", {"textFormat":{"bold":True}})

# performing calculations using formulas
sheet.update_cell(len(values)+1, 1, "Total")

# formatting the Total
sheet.format("A5", {"textFormat":{"bold":True}})

sheet.update_cell(len(values)+1, 2, "=sum(B2:B4)")  # sum of the prices
sheet.update_cell(len(values)+1, 3, "=sum(C2:C4)")  # sum of the quantities

for i in range(len(values)):
    print("row", i+1, ":", sheet.row_values(i+1))


    

