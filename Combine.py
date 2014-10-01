"""
Description:
Combine multiple CSV data files into one Excel document
"""

import os
import glob
import csv
import openpyxl
import natsort

filename = 'Test'

#Sets up workbook & adds data summary sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Data Summary'

#Gets list of all CSV files in the current directory and sorts them intelligently
files = glob.glob(os.path.join('.', '*.csv'))
sortedFiles = natsort.natsorted(files)

#Puts each CSV file into a new sheet
for csvfile in sortedFiles:
    head, tail = os.path.split(csvfile)
    ws = wb.create_sheet()
    ws.title = tail
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, val in enumerate(row):
                cell = ws.cell(row = r+1, column = c+1)
                try:
                    cell.value = float(val) #Trys to convert all numbers to floats
                except ValueError:
                    try:
                        cell.value = str(val) #Trys to submit the string value
                    except UnicodeDecodeError:
                        try:
                            print("Error decoding %s in cell (row = %i, column = %i)\n" \
                                  "Trying a different encoding standard" \
                                  %(val, r+1, c+1))
                            cell.value = unicode(val, "ISO-8859-1")
                            print("Different encoding standard worked!")
                        except:
                            print("Decoding failed\nThis cell will be blank")
                except:
                    print("Other Error\nError occured in cell "\
                          "(row = %i, column = %i). The value was %s"\
                          %(r,c,val))
                    print val
    print('Added %s' %tail)

wb.save(filename + '.xlsx')
print('Wrote ' + filename)
