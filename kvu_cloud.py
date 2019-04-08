# -*- coding: utf-8 -*-

import gspread
from oauth2client.service_account import ServiceAccountCredentials

old_link = "https://docs.google.com/spreadsheets/d/1rQbs3F-xBQC6Cf5zxlOfaDoU1ZkllUr4df595CoqRrU/edit?usp=sharing"
new_link = "https://docs.google.com/spreadsheets/d/1THkwfM4xqULwkw3E6Rm2JEoK78hWE4DladMZKlzJZXg/edit?usp=sharing"
old_filename = u"file:///Users/alekseydorofeev/kvu/Форма реестра.xlsx"
new_filename = u"file:///Users/alekseydorofeev/kvu/Форма реестра1.xlsx"
sheet_name = u"Таблица_ЗЛ"
new_sheet_name = u"Таблица_ЗЛ_изменения"

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

def found_changed_cols(row1, row2):
    cols = []
    i = 0
    for v in row1:
        if v != row2[i]:
            cols.append(i)
        i = i+1
    return cols

credentials = ServiceAccountCredentials.from_json_keyfile_name('kvu2019-6d3fd74effc3.json', scope)
gc = gspread.authorize(credentials)

db1 = gc.open(old_filename)
#db1 = gc.open_by_url(old_link)
wsh1 = db1.worksheet(sheet_name)
wsh1.duplicate(None, None, new_sheet_name)
wsh1 = db1.worksheet(new_sheet_name)

db2 = gc.open(new_filename)
#db2 = gc.open_by_url(new_link)
wsh2 = db2.worksheet(sheet_name)
wsh2.duplicate(None, None, new_sheet_name)
wsh2 = db2.worksheet(new_sheet_name)

val1 = wsh1.get_all_values()
val2 = wsh2.get_all_values()

i1 = 0
for row1 in val1:
    i1 = i1+1
    if i1 == 1:
        continue
    
    bill_num = row1[1]
    name = row1[4]
    typ = row1[5]
    passport = row1[6]
    uaddr = row1[7]
    paddr = row1[8]
    ao = row1[9]
    ag = row1[10]
    
    is_found = False
    is_changed = False
    changed_cols = []
    i2 = 0
    for row2 in val2:
        i2 = i2+1
        if i2 == 1:
            continue
        if bill_num == row2[1] and typ == row2[5]:
            if typ == u"ОС":
                is_found = True
                is_changed = (name!=row2[4] or ao!=row2[9] or ag!=row2[10])
            elif typ == u"ФЛ":
                if name == row2[4]:
                    is_found = True
                    is_changed = (passport!=row2[6] or uaddr!=row2[7] or paddr!=row2[8])
            elif typ == u"ЮЛ":
                is_found = True
                is_changed = (name!=row2[4] or uaddr!=row2[7] or paddr!=row2[8] or ao!=row2[9] or ag!=row2[10])
            if is_found:
                break

    if is_found:
        if is_changed:
            cols = found_changed_cols(row1, row2)
            print 'row #%d(%d): *' % (i1,i2)
            wsh1.update_cell(i1, 1, '***'+str(cols))
            wsh2.update_cell(i2, 1, '***'+str(cols))
        else:
            print 'row #%d(%d): v' % (i1,i2)
            wsh1.update_cell(i1, 1, 'vvv')
            wsh2.update_cell(i2, 1, 'vvv')
    else:
        print 'row #%d: -' % (i1)
        wsh1.update_cell(i1, 1, '---')
