"""Provide a class for DOCX  location descriptions export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2mso
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2msolib.docx.docx_file import OdtFile


class DocxLocations(OdtFile):
    """DOCX location descriptions file representation.

    Export a location sheet with  descriptions.
    """
    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _locationTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

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
