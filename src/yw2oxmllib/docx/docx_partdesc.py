"""Provide a class for DOCX  part descriptions export.

Parts are chapters marked `This chapter  begins a new section` in yWriter.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.docx.docx_file import DocxFile


class DocxPartDesc(DocxFile):
    """DOCX part summaries file representation.

    Export a synopsis with  part descriptions.
    """
    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _partTemplate = '''<w:p><w:pPr><w:pStyle w:val="heading1"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
