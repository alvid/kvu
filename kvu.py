# -*- coding: utf-8 -*-

import sys
import getopt
import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Fill
from openpyxl.styles import colors

verbose = 0
sheet_name = u"Таблица_ЗЛ"

# color for mark different records
rft = Font(color=colors.RED)
# color for mark missing records
bft = Font(color=colors.BLUE)

def found_changed_cols(row1, row2):
    cols = []
    i = 0
    for v in row1:
        if v.value != row2[i].value:
            cols.append(i)
        i = i+1
    return cols

def compare_files(filename1, filename2, new_filename):
    print "compare_files(%s, %s, %s)" % (filename1, filename2, new_filename)
    
    wb1 = load_workbook(filename=filename1)
    ws1 = wb1[sheet_name]

    wb2 = load_workbook(filename=filename2)
    ws2 = wb2[sheet_name]

    i1 = 1
    for row1 in ws1.rows:
        bill_num = row1[1].value
        name = row1[4].value
        typ = row1[5].value
        passport = row1[6].value
        uaddr = row1[7].value
        paddr = row1[8].value
        ao = row1[9].value
        ag = row1[10].value

	if bill_num is None and name is None and typ is None:
            print "end of input file is detected at row #%d" % (i1)
	    break
    
        is_found = False
        is_changed = False
        changed_cols = []
        i2 = 0
        for row2 in ws2.rows:
            i2 = i2+1
            if i2 == 1:
                continue
            if bill_num == row2[1].value and typ == row2[5].value:
                if typ == u"ОС":
                    is_found = True
                    is_changed = (name!=row2[4].value or ao!=row2[9].value or ag!=row2[10].value)
                elif typ == u"ФЛ":
                    if name == row2[4].value:
                        is_found = True
                        is_changed = (passport!=row2[6].value or uaddr!=row2[7].value or paddr!=row2[8].value or ao!=row2[9].value or ag!=row2[10].value)
                elif typ == u"ЮЛ":
                    is_found = True
                    is_changed = (name!=row2[4].value or uaddr!=row2[7].value or paddr!=row2[8].value or ao!=row2[9].value or ag!=row2[10].value)
                if is_found:
                    break
        # end of for row2 in ws2.rows:

        if is_found:
            if is_changed:
		if verbose:
                    print 'row #%d(%d): *' % (i1,i2)
                row1[1].font = rft
                row2[1].font = rft
                cols = found_changed_cols(row1, row2)
                for c in cols:
                    row1[c].font = rft
                    row2[c].font = rft
            else:
		if verbose:
                    print 'row #%d(%d): v' % (i1,i2)
        else:
	    if verbose:
                print 'row #%d: -' % (i1)
            for c in row1:
                c.font = bft

        i1 = i1+1
    # end of for row1 in ws1.rows:

    wb1.save(new_filename)

def usage(argv):
    print "Usage: " + argv[0] + " -o old_registry_filename -n new_registry_filename"
    print " or"
    print "Usage: " + argv[0] + " --old=old_registry_filename --new=new_registry_filename"

def main(argv=None):
    global verbose
    in_old_filename = u".\\files\\1.xlsx"
    out_old_filename = u".\\files\\1_diff.xlsx"
    in_new_filename = u".\\files\\2.xlsx"
    out_new_filename = u".\\files\\2_diff.xlsx"

    if argv is None:
        argv = sys.argv
    # Разбираем аргументы командной строки
    try:
        opts, args = getopt.getopt(argv[1:], "ho:n:v", ["help", "old=", "new=", "verbose"])
    except getopt.error, msg:
        print msg
        print "для справки используйте --help"
        sys.exit(1)
    #print >> sys.stderr, opts

    # Анализируем опции
    for o, arg in opts:
        if o in ("-h", "--help"):
            usage(argv)
            sys.exit(0)
        elif o in ("-v", "--verbose"):
            verbose = verbose + 1
        elif o in ("-o", "--old"):
            print "old registry file: " + arg
            in_old_filename = arg
            path=os.path.split(in_old_filename)
            if path[0]=='' or path[0]=='.':
	        out_old_filename = "diff_old.xlsx"
            else:
	        out_old_filename = path[0] + "\\diff_old.xlsx"
            print "difference between old and new registry will be saved on: " + out_old_filename
        elif o in ("-n", "--new"):
            print "new registry file: " + arg
            in_new_filename = arg
            path=os.path.split(in_new_filename)
            if path[0]=='' or path[0]=='.':
	        out_new_filename = "diff_new.xlsx"
            else:
	        out_new_filename = path[0] + "\\diff_new.xlsx"
            print "difference between new and old registry will be saved on: " + out_new_filename
    
    homepath = os.path.expanduser(os.getenv('USERPROFILE'))
    result = homepath + "\\kvu.result"

    try:
        compare_files(in_old_filename, in_new_filename, out_old_filename)
        compare_files(in_new_filename, in_old_filename, out_new_filename)
    except IOError as e:
	print "IOError happens, see %s" % result
	with open(result,"w") as f:
	    f.write(str(e))
        sys.exit(2)
    except Exception as e:
	print "UnknownError happens, see %s" % result
	with open(result,"w") as f:
	    f.write(str(e))
	sys.exit(3)
    else:
	print "work is successful completed"
	with open(result,"w") as f:
	    f.write("Normally completed")

if __name__ == "__main__":
    main()

