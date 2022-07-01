"""Provide a class for DOCX  character descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.docx.docx_file import DocxFile


class DocxCharacters(DocxFile):
    """DOCX character descriptions file representation.

    Export a character sheet with  descriptions.
    """
    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _characterTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title$FullName$AKA</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Description</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Bio</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Bio</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Goals</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Goals</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Notes</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Notes</w:t></w:r></w:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER

    def _get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section.
        
        Positional arguments:
            crId -- str: character ID.
        
        Special formatting of alternate and full name. 
        Extends the superclass method.
        """
        characterMapping = DocxFile._get_characterMapping(self, crId)
        if self.characters[crId].aka:
            characterMapping['AKA'] = f' ("{self.characters[crId].aka}")'
        if self.characters[crId].fullName:
            characterMapping['FullName'] = f'/{self.characters[crId].fullName}'
        return characterMapping
