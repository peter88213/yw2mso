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

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER
