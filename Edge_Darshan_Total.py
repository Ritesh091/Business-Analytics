import gspread
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import gspread_dataframe as gd
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('My Project 911-2177c70a5f89.json', scope)

gc = gspread.authorize(credentials)
wks = gc.open("Daily_Task_Inhouse")
a = wks.worksheet('Darshan Analysis')

filename = "total.xlsx"

book = openpyxl.load_workbook(filename)
writer = pd.ExcelWriter(filename, engine='openpyxl')

writer.book = book
writer.sheets = {x.title: x for x in book.worksheets}

test = a.get_all_records()
c = pd.DataFrame(test)
c = c[c.columns.drop(list(c.filter(regex='Unnamed')))]
print(c)

c.to_excel(writer, index=False)
writer.save()

g = wks.worksheet('Price')
jest = g.get_all_records()
r = pd.DataFrame(jest)
sal = r['Price'][0]

r.to_excel(writer, sheet_name='sheet4', index = False)
writer.save()

file = "ana1.xlsx"
s = pd.read_excel(file, header=None, index_col=False, skiprows = 1)

av_prce = (sal/25)/8
print(av_prce)

show = pd.DataFrame()
print(av_prce)
show['Darshan'] = av_prce

show.loc[0] = [av_prce]

c1 = c*av_prce

c1.to_excel(writer ,header= None, index=False, startrow = 2)
writer.save()

filename1 = "proj_total.xlsx"

book1 = openpyxl.load_workbook(filename1)
writer1 = pd.ExcelWriter(filename1, engine='openpyxl')

c1.to_excel(writer1 ,  index = False)
writer1.save()

ba = pd.read_excel(filename)

try:
    ap = wks.worksheet("Darshan Price")
    wks.del_worksheet(ap)
except:
    print('none')

try:
    worksheet = wks.add_worksheet(title="Darshan Price", rows="1", cols="2")
except:
    worksheet = gc.open("Daily_Task_Inhouse").worksheet("Darshan Price")
wks.del_worksheet(a)

existing = gd.get_as_dataframe(worksheet)
updated = existing.append(ba)
gd.set_with_dataframe(worksheet, updated)
