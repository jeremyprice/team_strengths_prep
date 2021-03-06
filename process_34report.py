#!/usr/bin/env python3

import os
import os.path
import fitz


def process(r34_fname, r34_dir, p21_dir, base_fname):
    '''process the 34 report - do all the things we need with this file:
    split the file into the individual reports
    extract page 21 and join the pages into a single report
    '''
    split_and_extract(r34_fname, r34_dir, p21_dir, base_fname)

def find_name(page, no_spaces=True):
    blocks = page.getText('blocks')
    # the name block looks like this: FIRST LAST | DATE
    name_blocks = [b for b in blocks if ' | ' in b[4]]
    if len(name_blocks) != 1:
        print('error: too many or too few name blocks: {}'.format(name_blocks))
        return 'not_found'
    # returned tuple: (x0, y0, x1, y1, "word", block_no, line_no, word_no)
    name = name_blocks[0][4]  # grab the only item we found, the 4th item in the tuple
    name = name.split('|')[0].strip()  # grab the name from the formatting and strip the whitespace
    name = name.title()
    if no_spaces:
        name = name.replace(' ', '')
    return name

def check_for_caution(page):
    blocks = page.getText('blocks')
    # the first page of the report contains A NOTE OF CAUTION
    name_blocks = [b for b in blocks if 'A NOTE OF CAUTION' in b[4]]
    # if there is a block with this text, we have a caution page
    return len(name_blocks) == 1

def split_and_extract(fname, r34_outdir, p21_outdir, base_fname):
    '''Split a Gallup full 34 report into the individual files
    @arg fname: name of the file to split
    @arg r34_outdir: the output directory to put all the 34 report split files into
    @arg p21_outdir: the output directory to put all the p21 extracted files into
    @arg base_fname: the base filename to prepend to any generated files {}-p21.pdf
    '''
    doc = fitz.open(fname)
    numpages = doc.pageCount
    report_count = numpages // 26
    if numpages % 26 != 0:
        report_count += 1
    fname_34fmt = os.path.join(r34_outdir, '{}-34report.pdf')
    fname_21fmt = os.path.join(p21_outdir, '{}-p21.pdf')
    p21_all = []
    for report_num in range(report_count):
        start_page_num = (report_num * 26)
        stop_page_num = start_page_num + 25  # we only need the first 25 pages since the last is blank
        p21 = 20
        person_name = find_name(doc[start_page_num])
        if check_for_caution(doc[start_page_num]):
            stop_page_num += 1
            p21 += 1
        r34_fname = fname_34fmt.format(person_name)
        p21_fname = fname_21fmt.format(person_name)
        doc.select(range(start_page_num, stop_page_num))
        doc.save(r34_fname, garbage=4, clean=1, deflate=1)
        doc.select([p21])
        doc.save(p21_fname, garbage=4, clean=1, deflate=1)
        doc.close()
        doc = fitz.open(fname)
        # hang on to the person's name and p21 index so we can save it to it's own file
        p21_all.append((person_name, start_page_num + p21))
    # save all the p21s to a single file
    p21_all.sort()  # sort the pages by first name alphabetical
    doc.select([p[1] for p in p21_all])
    p21_fname = '{}-p21.pdf'.format(base_fname)
    doc.save(p21_fname, garbage=4, clean=1, deflate=1)


if __name__ == '__main__':
    import sys
    process(*sys.argv[1:])
