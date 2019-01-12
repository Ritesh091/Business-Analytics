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
a = wks.get_worksheet(1)

test = a.get_all_records()

df = pd.DataFrame.from_records(test[2:], columns=test[0])
print(df.columns.astype(int))

df.to_excel("ana1.xlsx")

file = "ana1.xlsx"
s = pd.read_excel(file, header=None, index_col=False, skiprows = 1)
g = len(s.columns)
print(g)

f = []
l = []
cont = {col: s[col] for col in s.columns}
for i in range(2,g):
    c = cont[i]
    d = c.value_counts()
    r = d*.5
    l.append(r)
    j = c[0]
    f.append(j)

pet = pd.DataFrame(l).stack()

book = openpyxl.load_workbook(file)
writer = pd.ExcelWriter(file, engine='openpyxl')

writer.book = book
writer.sheets = {x.title: x for x in book.worksheets}

pet.to_excel(writer, sheet_name='sheet2')
writer.save()

sheet = writer.book.get_sheet_by_name('sheet2')
sheet['B1'] = str('TASK')
writer.save()

xls = pd.ExcelFile(file)
df1 = pd.read_excel(xls, 'sheet2', index_col=0)

p = df1.groupby('TASK')[0].sum()

p.to_excel(writer, sheet_name='sheet3')
writer.save()

show = pd.DataFrame()
try:
    a = [p[p.index[pd.Series(p.index).str.startswith('REMP')]]]
    jet = pd.DataFrame(a).stack()
    jets = jet[0].sum()
    show['REMP'] = [jets]
except:
    show['REMP'] = [0]

try:
    b = [p[p.index[pd.Series(p.index).str.startswith('AARON')]]]
    pet = pd.DataFrame(b).stack()
    pets = pet[0].sum()
    show['AARON'] = [pets]
except:
    show['AARON'] = [0]

try:
    c = [p[p.index[pd.Series(p.index).str.startswith('DAGHA')]]]
    let = pd.DataFrame(c).stack()
    lets = let[0].sum()
    show['DAGHA'] = [lets]
except:
    show['DAGHA'] = [0]

try:
    d = [p[p.index[pd.Series(p.index).str.startswith('EDGE')]]]
    set = pd.DataFrame(d).stack()
    sets = set[0].sum()
    show['EDGE'] = [sets]
except:
    show['EDGE'] = [0]

try:
    e = [p[p.index[pd.Series(p.index).str.startswith('NRI')]]]
    tet = pd.DataFrame(e).stack()
    tets = tet[0].sum()
    show['NRI'] = [tets]
except:
    show['NRI'] = [0]

try:
    f = [p[p.index[pd.Series(p.index).str.startswith('STARTUPP')]]]
    met = pd.DataFrame(f).stack()
    mets = met[0].sum()
    show['STARTUPP'] = [mets]
except:
    show['STARTUPP'] = [0]

try:
    h = [p[p.index[pd.Series(p.index).str.startswith('ICO')]]]
    ret = pd.DataFrame(h).stack()
    rets = ret[0].sum()
    show['ICO'] = [rets]
except:
    show['ICO'] = [0]

try:
    i = [p[p.index[pd.Series(p.index).str.startswith('LAB')]]]
    wet = pd.DataFrame(i).stack()
    wets = wet[0].sum()
    show['LAB'] = [wets]
except:
    show['LAB'] = [0]

try:
    j = [p[p.index[pd.Series(p.index).str.startswith('BLOCK')]]]
    ket = pd.DataFrame(j).stack()
    kets = ket[0].sum()
    show['BLOCK'] = [kets]
except:
    show['BLOCK'] = [0]

try:
    k = [p[p.index[pd.Series(p.index).str.startswith('ACME')]]]
    yet = pd.DataFrame(k).stack()
    yets = yet[0].sum()
    show['ACME'] = [yets]
except:
    show['ACME'] = [0]

try:
    l = [p[p.index[pd.Series(p.index).str.startswith('AWS')]]]
    fret = pd.DataFrame(l).stack()
    frets = fret[0].sum()
    show['AWS'] = [frets]
except:
    show['AWS'] = [0]

try:
    m = [p[p.index[pd.Series(p.index).str.startswith('R&D')]]]
    phet = pd.DataFrame(m).stack()
    phets = phet[0].sum()
    show['R&D'] = [phets]
except:
    show['R&D'] = [0]

show.plot(kind="bar",  figsize=(10,5))
plt.ylabel('Hours')
plt.xlabel('Tasks')
plt.show()

try:
    ap = wks.worksheet("Darshan Analysis")
    wks.del_worksheet(ap)
except:
    print('none')

try:
    worksheet = wks.add_worksheet(title="Darshan Analysis", rows="1", cols="2")
except:
    worksheet = gc.open("Daily_Task_Inhouse").worksheet("Darshan Analysis")

existing = gd.get_as_dataframe(worksheet)
updated = existing.append(show)
gd.set_with_dataframe(worksheet, updated)

