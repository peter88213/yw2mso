"""Provide a class for DOCX brief synopsis export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import OdtFile


class DocxBriefSynopsis(OdtFile):
    """DOCX brief synopsis file representation.

    Export a brief synopsis with chapter titles and scene titles.
    """
    DESCRIPTION = 'Brief synopsis'
    SUFFIX = '_brf_synopsis'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = '''<text:p text:style-name="Text_20_body">$Title</text:p>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER
