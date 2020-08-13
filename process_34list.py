#!/usr/bin/env python3

from operator import itemgetter
import sys
import xlrd
import xlsxwriter

def process(fname, base_fname):
    '''process the given 34 list Excel file and generate the name tents and the team matrix
    @arg fname: the 34 list Excel filename
    @arg base_fname: the base filename string to prepend to any generated files {}-matrix.xlsx
    '''
    info = load_info(fname, base_fname)

def save_34list(info, base_fname):
    '''save the updated and sorted data to a new Excel file'''
    fname = '{}-34list.xlsx'.format(base_fname)
    wb = xlsxwriter.Workbook(fname)
    ws = wb.add_worksheet()
    # write the header row
    ws.write(0, 0, 'Name (Last, First)')
    ws.write(0, 1, 'Name')
    for theme in range(1, 35):
        ws.write(0, theme+1, 'Theme {}'.format(theme))
    for row, person in enumerate(info):
        for col in range(0, len(person)):
            ws.write(row+1, col, person[col])
    # add in the autofilter drop downs
    ws.autofilter(0, 0, len(info), 35)
    wb.close()

def first_last_name(name_lf):
    '''split a last, first name into first last order'''
    parts = name_lf.split(',')
    parts.reverse()
    name = ' '.join(parts)
    return name.title().strip()

def load_info(fname, base_fname):
    '''load the existing information from the Gallup generated Excel file
    returns the information as a dictionary keyed by person's name and the value is
    a list of the 34 strengths in order'''
    wb = xlrd.open_workbook(fname)
    ws = wb.sheet_by_index(0)
    first_row = True
    info = []
    for row in ws.get_rows():
        if first_row:  # create the map keys from the header row
            header = {cell.value:idx for idx, cell in enumerate(row)}
            if 'Theme 34' in header:  # check if this is a top 5 or full 34
                last_col = 35
            else:
                last_col = 6
            first_row = False
            continue
        namelf = row[header['Name (Last, First)']].value
        if namelf == '':
            # this is the end of the data
            break
        namefl = first_last_name(namelf)
        person = [namelf, namefl, *[row[header['Theme {}'.format(idx)]].value for idx in range(1,last_col)]]
        info.append(person)
    # sort it by first name
    info.sort(key=itemgetter(1))
    save_34list(info, base_fname)

if __name__ == '__main__':
    process(*sys.argv[1:])
