"""Provide a class for DOCX item  descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.docx.docx_file import DocxFile


class DocxItems(DocxFile):
    """DOCX item descriptions file representation.

    Export a item sheet with  descriptions.
    """
    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _itemTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title$AKA</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER

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
