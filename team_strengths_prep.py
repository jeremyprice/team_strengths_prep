#!/usr/bin/env python3

import os
import os.path
import shutil
import sys
import tempfile
import process_34report
import process_34list
import nametents

def prep_staging(staging_dir, r34_dir, p21_dir, r34_fname, list34_fname):
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
    fname = os.path.basename(list34_fname)
    new_list34_fname = os.path.join(staging_dir, fname)
    shutil.copyfile(list34_fname, new_list34_fname)
    nametents.prep_staging(staging_dir)

def main():
    r34_fname = sys.argv[1]
    list34_fname = sys.argv[2]
    base_fname = sys.argv[3]
    r34_dir = '34reports'
    p21_dir = 'p21'
    # TODO: use a temporary staging dir so it will clean up the files when we're done
    # staging_dir = tempfile.TemporaryDirectory()
    # staging_dir_name = staging_dir.name
    staging_dir_name = 'staging'
    prep_staging(staging_dir_name, r34_dir, p21_dir, r34_fname, list34_fname)
    prev_dir = os.getcwd()
    os.chdir(staging_dir_name)
    r34_fname = os.path.basename(r34_fname)
    list34_fname = os.path.basename(list34_fname)
    process_34report.process(r34_fname, r34_dir, p21_dir, base_fname)
    process_34list.process(list34_fname, base_fname)
    # if everything went well we can get rid of the xls 34list file
    os.remove(list34_fname)
    os.chdir(prev_dir)

if __name__ == '__main__':
    main()
