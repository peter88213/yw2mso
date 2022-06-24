"""Provide a class for DOCX item  descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import OdtFile


class DocxItems(OdtFile):
    """DOCX item descriptions file representation.

    Export a item sheet with  descriptions.
    """
    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _itemTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section.
        
        Positional arguments:
            itId -- str: item ID.
        
        Special formatting of alternate name. 
        Extends the superclass method.
        """
        itemMapping = super()._get_itemMapping(itId)
        if self.items[itId].aka:
            itemMapping['AKA'] = f' ("{self.items[itId].aka}")'
        return itemMapping
