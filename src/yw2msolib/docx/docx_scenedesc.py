"""Provide a class for DOCX  scene descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import DocxFile


class DocxSceneDesc(DocxFile):
    """DOCX scene summaries file representation.

    Export a full synopsis with  scene descriptions.
    """
    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _partTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _chapterTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _sceneTemplate = '''<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>
'''

    _sceneDivider = '<w:p><w:pPr><w:pStyle w:val="Heading4"/><w:ind w:hanging="0"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>* * *</w:t></w:r></w:p>\n'
    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
