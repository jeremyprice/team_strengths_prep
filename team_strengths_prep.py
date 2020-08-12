#!/usr/bin/env python3

import os
import os.path
import shutil
import sys
import tempfile
import process_34report

def prep_34report(staging_dir, r34_dir, p21_dir, r34_fname):
    for d in (r34_dir, p21_dir):
        newdir = os.path.join(staging_dir, d)
        try:
            os.mkdir(newdir)
        except FileExistsError:
            # clear the directory if it already exists
            shutil.rmtree(newdir, ignore_errors=True)
            os.mkdir(newdir)
    fname = os.path.basename(r34_fname)
    new_r34_fname = os.path.join(staging_dir, fname)
    shutil.copyfile(r34_fname, new_r34_fname)

def main():
    r34_fname = sys.argv[1]
    r34_dir = '34reports'
    p21_dir = 'p21'
    # TODO: use a temporary staging dir so it will clean up the files when we're done
    # staging_dir = tempfile.TemporaryDirectory()
    # staging_dir_name = staging_dir.name
    staging_dir_name = 'staging'
    prep_34report(staging_dir_name, r34_dir, p21_dir, r34_fname)
    prev_dir = os.getcwd()
    os.chdir(staging_dir_name)
    r34_fname = os.path.basename(r34_fname)
    process_34report.process(r34_fname, r34_dir, p21_dir)
    os.chdir(prev_dir)

if __name__ == '__main__':
    main()
