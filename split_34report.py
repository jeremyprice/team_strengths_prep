#!/usr/bin/env python3

from PyPDF2 import PdfFileReader, PdfFileWriter
import subprocess


def split(fname, outdir, parallel=True):
    '''Split a Gallup full 34 report into the individual files
    @arg fname: name of the file to split
    @arg outdir: the output directory to put all the split files into
    '''
    reader = PdfFileReader(fname)
    report_count = reader.getNumPages() // 26
    if reader.getNumPages() % 26 != 0:
        report_count += 1
    if report_count > 10:
        fname_fmt = '{}/{:02d}.pdf'
    elif report_count > 100:
        fname_fmt = '{}/{:03d}.pdf'
    else:
        fname_fmt = '{}/{:1d}.pdf'
    with open('gs_input', 'wb') as outfile:
        outfile.write(b'\n' * 1000)
    subproc_stdin = open('gs_input', 'rb')
    procs = []
    for report_num in range(report_count):
        start_page_num = (report_num * 26) + 1
        stop_page_num = start_page_num + 25
        if stop_page_num > reader.getNumPages():
            stop_page_num = reader.getNumPages()
        out_fname = fname_fmt.format(outdir, report_num)
        print("Page {}: {}".format(start_page_num, report_num))
        if parallel:
            proc = subprocess.Popen(["gs", "-dBATCH",  '-sOutputFile={}'.format(out_fname),
                                     "-dFirstPage={}".format(start_page_num),
                                     "-dLastPage={}".format(stop_page_num),
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

if __name__ == '__main__':
    import sys
    split(sys.argv[1], sys.argv[2])
