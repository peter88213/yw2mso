"""Regression test for the yw2oxml project.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import unittest
from shutil import copyfile, rmtree, copytree
import re
import zipfile

from yw2oxmllib.yw2oxml_exporter import Yw2msoExporter

UPDATE = False

# Test environment

# The paths are relative to the "test" directory,
# where this script is placed and executed

TEST_PATH = f'{os.getcwd()}/../test'
TEST_DATA_PATH = f'{TEST_PATH}/data/'
TEST_EXEC_PATH = f'{TEST_PATH}/'

# To be placed in TEST_DATA_PATH:
YW7_NORMAL = 'normal.yw7'
DOCX_NORMAL = 'normal.docx'
DOCUMENT_NORMAL_EXPORT = 'export.xml'

# Test data
DOCUMENT = 'document.xml'
PROJECT = 'Sample Project'


def read_file(inputFile):
    try:
        with open(inputFile, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        # HTML files exported by a word processor may be ANSI encoded.
        with open(inputFile, 'r') as f:
            return f.read()


def remove_all_testfiles():
    try:
        os.remove(f'{TEST_EXEC_PATH}{PROJECT}.yw7')
    except:
        pass
    try:
        os.remove(f'{TEST_EXEC_PATH}{PROJECT}.docx')
    except:
        pass
    try:
        rmtree(f'{TEST_EXEC_PATH}word')
    except:
        pass


class NormalOperation(unittest.TestCase):
    """Test case: Normal operation."""

    def setUp(self):
        try:
            os.mkdir(TEST_EXEC_PATH)
        except:
            pass
        remove_all_testfiles()
        self.exporter = Yw2msoExporter()
        self.exporter.ui = UiStub()

    def test_yw7_to_docx(self):
        copyfile(f'{TEST_DATA_PATH}{YW7_NORMAL}', f'{TEST_EXEC_PATH}{PROJECT}.yw7')
        os.chdir(TEST_EXEC_PATH)
        kwargs = {'suffix': ''}
        self.exporter.run(f'{TEST_EXEC_PATH}{PROJECT}.yw7', **kwargs)
        with zipfile.ZipFile(f'{TEST_EXEC_PATH}{PROJECT}.docx', 'r') as myzip:
            myzip.extract(f'word/{DOCUMENT}', TEST_EXEC_PATH)
        if UPDATE:
            copyfile(f'word/{DOCUMENT}', f'{TEST_DATA_PATH}{DOCUMENT_NORMAL_EXPORT}')
        self.assertEqual(read_file(f'word/{DOCUMENT}'), read_file(f'{TEST_DATA_PATH}{DOCUMENT_NORMAL_EXPORT}'))

    def tearDown(self):
        remove_all_testfiles()


class UiStub:

    def show_open_button(self):
        return

    def set_info_what(self, message):
        return

    def set_info_how(self, message):
        return

    def ask_yes_no(self, text):
        return True


def main():
    unittest.main()


if __name__ == '__main__':
    main()
