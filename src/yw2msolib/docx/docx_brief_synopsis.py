"""Provide a class for DOCX brief synopsis export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import DocxFile


class DocxBriefSynopsis(DocxFile):
    """DOCX brief synopsis file representation.

    Export a brief synopsis with chapter titles and scene titles.
    """
    DESCRIPTION = 'Brief synopsis'
    SUFFIX = '_brf_synopsis'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _partTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _chapterTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _sceneTemplate = '''<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Title</w:t></w:r></w:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
