"""
Description:
Combine multiple CSV data files into one Excel document
"""

import os
import glob
import csv
import xlwt
import natsort

filename = 'Test'

wb = xlwt.Workbook()
ws = wb.add_sheet('Data Summary')

files = glob.glob(os.path.join('.', '*.csv'))
sortedFiles = natsort.natsorted(files)

for csvfile in sortedFiles:
    head, tail = os.path.split(csvfile)
    ws = wb.add_sheet(tail)
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, val in enumerate(row):
                if r > 4:
                    ws.write(r,c,float(val)) #Cconverts all data to floats
                else:
                    ws.write(r, c, val)
    print 'Added %s' %csvfile

wb.save(filename + '.xlsx')
print 'Wrote ' + filename
