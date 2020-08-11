#!/usr/bin/env python3

import.os
import os.path
import sys
import tempfile
import process_34report

def prep_34report(staging_dir, r34_dir, p21_dir):
    [os.mkdir(os.path.join(staging_dir, d)) for d in (r34_dir, p21_dir)]

def main():
    r34_fname = sys.argv[1]
    r34_dir = '34reports'
    p21_dir = 'p21'
    staging_dir = tempfile.TemporaryDirectory()
    print(staging_dir)
    prep_34report(staging_dir, r34_dir, p21_dir)
    # process_34report.process()

if __name__ == '__main__':
    main()
