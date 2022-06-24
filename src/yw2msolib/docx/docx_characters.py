"""Provide a class for DOCX  character descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import OdtFile


class DocxCharacters(OdtFile):
    """DOCX character descriptions file representation.

    Export a character sheet with  descriptions.
    """
    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _characterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$FullName$AKA</text:h>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Description</text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Bio</text:h>
<text:p text:style-name="Text_20_body">$Bio</text:p>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Goals</text:h>
<text:p text:style-name="Text_20_body">$Goals</text:p>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Notes</text:h>
<text:p text:style-name="Text_20_body">$Notes</text:p>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section.
        
        Positional arguments:
            crId -- str: character ID.
        
        Special formatting of alternate and full name. 
        Extends the superclass method.
        """
        characterMapping = OdtFile._get_characterMapping(self, crId)
        if self.characters[crId].aka:
            characterMapping['AKA'] = f' ("{self.characters[crId].aka}")'
        if self.characters[crId].fullName:
            characterMapping['FullName'] = f'/{self.characters[crId].fullName}'
        return characterMapping
