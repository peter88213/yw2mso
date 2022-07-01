"""Provide a generic class for Open XML file export.

All XLSX and DOCX file representations inherit from this class.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import zipfile
import tempfile
from shutil import rmtree
import locale
from datetime import datetime
from string import Template

from pywriter.pywriter_globals import ERROR
from pywriter.file.file_export import FileExport


class OxmlFile(FileExport):
    """Generic Open XML file representation.

    Public methods:
        write() -- write instance variables to the export file.
    """
    _OXML_COMPONENTS = []
    _CONTENT_TYPES_XML = ''
    _CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dcterms:created xsi:type="dcterms:W3CDTF">$Datetime</dcterms:created>
    <dc:creator>$Author</dc:creator>
    <dc:description>$Summary</dc:description>
    <cp:keywords>  </cp:keywords>
    <dc:language>$Language-$Country</dc:language>
    <dc:title>$Title</dc:title>
</cp:coreProperties>
'''
    _CUSTOM_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>
'''
    _RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>  
'''

    def __init__(self, filePath, **kwargs):
        """Create a temporary directory for zipfile generation.
        
        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        Extends the superclass constructor,        
        """
        super().__init__(filePath, **kwargs)
        self._tempDir = tempfile.mkdtemp(suffix='.tmp', prefix='oxml_')
        self._originalPath = self._filePath

    def __del__(self):
        """Make sure to delete the temporary directory, in case write() has not been called."""
        self._tear_down()

    def _tear_down(self):
        """Delete the temporary directory containing the unpacked OXML directory structure."""
        try:
            rmtree(self._tempDir)
        except:
            pass

    def _set_up(self):
        """Helper method for ZIP file generation.

        Prepare the temporary directory containing the internal structure of an OXML file except 'content.xml'.
        Return a message beginning with the ERROR constant in case of error.
        """

        #--- Create and open a temporary directory for the files to zip.
        try:
            self._tear_down()
            os.mkdir(self._tempDir)
            os.mkdir(f'{self._tempDir}/_rels')
            os.mkdir(f'{self._tempDir}/docProps')
        except:
            return f'{ERROR}Cannot create "{os.path.normpath(self._tempDir)}".'

        #--- Generate [Content_Types].xml.
        try:
            with open(f'{self._tempDir}/[Content_Types].xml', 'w', encoding='utf-8') as f:
                f.write(self._CONTENT_TYPES_XML)
        except:
            return f'{ERROR}Cannot write "[Content_Types].xml"'

        #--- Generate docProps/core.xml.

        #  Set system language set as document language.
        lng, ctr = locale.getdefaultlocale()[0].split('_')
        coreMapping = dict(
            Author=self.authorName,
            Title=self.title,
            Summary=f'<![CDATA[{self.desc}]]>',
            Datetime=datetime.today().replace(microsecond=0).isoformat(),
            Language=lng,
            Country=ctr,
       )
        template = Template(self._CORE_XML)
        text = template.safe_substitute(coreMapping)
        try:
            with open(f'{self._tempDir}/docProps/core.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Cannot write "core.xml".'

        #--- Generate docProps/custom.xml.
        try:
            with open(f'{self._tempDir}/docProps/custom.xml', 'w', encoding='utf-8') as f:
                f.write(self._CUSTOM_XML)
        except:
            return f'{ERROR}Cannot write "custom.xml"'

        #--- Generate _rels\.rels.
        try:
            with open(f'{self._tempDir}/_rels/.rels', 'w', encoding='utf-8') as f:
                f.write(self._RELS)
        except:
            return f'{ERROR}Cannot write ".rels"'

        return 'OXML structure generated.'

    def write(self):
        """Write instance variables to the export file.
        
        Create a template-based output file. 
        Return a message beginning with the ERROR constant in case of error.
        Extends the super class method, adding ZIP file operations.
        """

        #--- Create a temporary directory
        # containing the internal structure of an XLSX file except "content.xml".
        message = self._set_up()
        if message.startswith(ERROR):
            return message

        #--- Pack the contents of the temporary directory into the OXML file.
        workdir = os.getcwd()
        backedUp = False
        if os.path.isfile(self.filePath):
            try:
                os.replace(self.filePath, f'{self.filePath}.bak')
                backedUp = True
            except:
                return f'{ERROR}Cannot overwrite "{os.path.normpath(self.filePath)}".'

        try:
            with zipfile.ZipFile(self.filePath, 'w') as oxmlTarget:
                os.chdir(self._tempDir)
                for file in self._OXML_COMPONENTS:
                    oxmlTarget.write(file, compress_type=zipfile.ZIP_DEFLATED)
        except:
            os.chdir(workdir)
            if backedUp:
                os.replace(f'{self.filePath}.bak', self.filePath)
            return f'{ERROR}Cannot generate "{os.path.normpath(self.filePath)}".'

        #--- Remove temporary data.
        os.chdir(workdir)
        self._tear_down()
        return f'"{os.path.normpath(self.filePath)}" written.'
