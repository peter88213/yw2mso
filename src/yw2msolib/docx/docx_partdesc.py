"""Provide a class for DOCX  part descriptions export.

Parts are chapters marked `This chapter  begins a new section` in yWriter.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import DocxFile


class DocxPartDesc(DocxFile):
    """DOCX part summaries file representation.

    Export a synopsis with  part descriptions.
    """
    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
