"""Provide a class for DOCX  location descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.docx.docx_file import DocxFile


class DocxLocations(DocxFile):
    """DOCX location descriptions file representation.

    Export a location sheet with  descriptions.
    """
    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    _fileHeader = f'''{DocxFile._DOCUMENT_XML_HEADER}<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>$Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>$AuthorName</w:t></w:r></w:p>
'''

    _locationTemplate = '''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$Title$AKA</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>$Desc</w:t></w:r></w:p>
'''

    _fileFooter = DocxFile._DOCUMENT_XML_FOOTER

    def _get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section.
        
        Positional arguments:
            lcId -- str: location ID.
        
        Special formatting of alternate name. 
        Extends the superclass method.
        """
        locationMapping = super()._get_locationMapping(lcId)
        if self.locations[lcId].aka:
            locationMapping['AKA'] = f' ("{self.locations[lcId].aka}")'
        return locationMapping
