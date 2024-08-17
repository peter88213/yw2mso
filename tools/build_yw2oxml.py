"""Build a Python script for the OpenOffice "convert yWriter" script.
        
Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
sys.path.insert(0, f'{os.getcwd()}/../../yw2oxml/src')
import inliner

SRC = '../src/'
BUILD = '../test/'
SOURCE_FILE = f'{SRC}yw2oxml_.pyw'
TARGET_FILE = f'{BUILD}yw2oxml.pyw'


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'yw2oxmllib', '../src/', copyPyWriter=True)
    inliner.run(TARGET_FILE, TARGET_FILE, 'pywriter', '../../PyWriter/src/', copyPyWriter=True)
    # inliner.run(SOURCE_FILE, TARGET_FILE, 'yw2oxmllib', '../src/')
    # inliner.run(TARGET_FILE, TARGET_FILE, 'pywriter', '../src/')
    print('Done.')


if __name__ == '__main__':
    main()
