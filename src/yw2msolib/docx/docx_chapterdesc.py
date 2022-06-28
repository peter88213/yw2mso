"""Provide a class for DOCX  chapter descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import DocxFile


class DocxChapterDesc(DocxFile):
    """DOCX chapter summaries file representation.

    Export a synopsis with  chapter descriptions.
    """
    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _partTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _chapterTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
