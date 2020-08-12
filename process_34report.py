#!/usr/bin/env python3

from PyPDF2 import PdfFileReader, PdfFileWriter
import subprocess
import os
import fitz


# TODO: rename the extracted files based on the person's name
# TODO: handle the banner of shame for p21 extraction


def process(r34_fname, r34_dir, p21_dir):
    '''process the 34 report - do all the things we need with this file:
    split the file into the individual reports
    extract page 21 and join the pages into a single report
    '''
    split_and_extract(r34_fname, r34_dir, p21_dir)

def find_name(doc, page_num, no_spaces=True):
    page = doc[page_num]
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

def split_and_extract(fname, r34_outdir, p21_outdir, parallel=True):
    '''Split a Gallup full 34 report into the individual files
    @arg fname: name of the file to split
    @arg outdir: the output directory to put all the split files into
    '''
    reader = PdfFileReader(fname)
    doc = fitz.open(fname)
    numpages = reader.getNumPages()
    report_count = numpages // 26
    if numpages % 26 != 0:
        report_count += 1
    fname_34fmt = os.path.join(r34_outdir, '{}-34report.pdf')
    fname_21fmt = os.path.join(p21_outdir, '{}-p21.pdf')
    with open('gs_input', 'wb') as outfile:
        outfile.write(b'\n' * 1000)
    subproc_stdin = open('gs_input', 'rb')
    procs = []
    for report_num in range(report_count):
        start_page_num = (report_num * 26) + 1
        stop_page_num = start_page_num + 25
        if stop_page_num > numpages:
            stop_page_num = numpages
        p21 = (report_num * 26) + 1 + 20
        person_name = find_name(doc, start_page_num)
        print(person_name)
        r34_fname = fname_34fmt.format(person_name)
        p21_fname = fname_21fmt.format(person_name)
        print("Page {}: {}".format(start_page_num, report_num))
        if parallel:
            proc = subprocess.Popen(["gs", "-dBATCH",  '-sOutputFile={}'.format(r34_fname),
                                     "-dFirstPage={}".format(start_page_num),
                                     "-dLastPage={}".format(stop_page_num),
                                     "-sDEVICE=pdfwrite", fname], stdin=subproc_stdin,
                                     stdout=subprocess.DEVNULL)
            procs.append(proc)
            proc = subprocess.Popen(["gs", "-dBATCH",  '-sOutputFile={}'.format(p21_fname),
                                     "-dFirstPage={}".format(p21),
                                     "-dLastPage={}".format(p21),
                                     "-sDEVICE=pdfwrite", fname], stdin=subproc_stdin,
                                     stdout=subprocess.DEVNULL)
            procs.append(proc)
        else:
            subprocess.run(["gs", "-dBATCH",  '-sOutputFile={}'.format(out_fname),
                            "-dFirstPage={}".format(start_page_num),
                            "-dLastPage={}".format(stop_page_num),
                            "-sDEVICE=pdfwrite", fname], stdin=subproc_stdin,
                            stdout=subprocess.DEVNULL)
    for proc in procs:
        proc.wait()
    os.remove('gs_input')

if __name__ == '__main__':
    import sys
    split(sys.argv[1], sys.argv[2])
