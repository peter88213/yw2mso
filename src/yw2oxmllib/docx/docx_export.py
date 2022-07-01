"""Provide a class for DOCX chapters and scenes export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.docx.docx_file import DocxFile


class DocxExport(DocxFile):
    """DOCX novel file representation.

    Export a non-reimportable manuscript with chapters and scenes.
    """
    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _partTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _chapterTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
'''

    _sceneTemplate = '''<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>$SceneContent</w:t></w:r></w:p>
'''

    _appendedSceneTemplate = '''<w:p><w:pPr><w:pStyle w:val="BodyTextIndent"/></w:pPr><w:r><w:t>$SceneContent</w:t></w:r></w:p>
'''

    _sceneDivider = '<w:p><w:pPr><w:pStyle w:val="Heading4"/><w:ind w:hanging="0"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>* * *</w:t></w:r></w:p>\n'
    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER

    def _get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section.
        
        Positional arguments:
            chId -- str: chapter ID.
            chapterNumber -- int: chapter number.
        
        Suppress the chapter title if necessary.
        Extends the superclass method.
        """
        chapterMapping = super()._get_chapterMapping(chId, chapterNumber)
        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''
        return chapterMapping
