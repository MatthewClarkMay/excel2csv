#!/usr/bin/python

import csv
import xlrd
from os import sys, path, makedirs


def excel_to_csv(excel_file):
    workbook = xlrd.open_workbook(excel_file)
    all_worksheets = workbook.sheet_names()
    for worksheet_name in all_worksheets:
        worksheet = workbook.sheet_by_name(worksheet_name)
        with open('convert_outdir/{}.csv'.format(excel_file + '_' + worksheet_name + '_conversion'), 'wb') as csv_file:
            wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
            for rownum in xrange(worksheet.nrows):
                wr.writerow([unicode(entry).encode('utf-8') for entry in worksheet.row_values(rownum)])


def make_out_dir():
    if not path.isdir('convert_outdir'):
      try:
          makedirs('convert_outdir')
      except OSError:
          print('ERROR: No convert_outdir created')
    else:
        print('Directory already exists!')


def main():
    make_out_dir()
    if path.isdir('convert_outdir'):
        for argcount in range(len(sys.argv)-1):
            excel_to_csv(sys.argv[argcount+1])
    else:
        print('Directory does not exist and it could not be created...')
        sys.exit()
    

if __name__ == '__main__':
    main()
