"""Provide a yWriter to MS Office converter.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yW2OO
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.yw_cnv_ff import YwCnvFf
from pywriter.yw.yw7_file import Yw7File

from yw2oxmllib.docx.docx_scenedesc import DocxSceneDesc
from yw2oxmllib.docx.docx_chapterdesc import DocxChapterDesc
from yw2oxmllib.docx.docx_partdesc import DocxPartDesc
from yw2oxmllib.docx.docx_brief_synopsis import DocxBriefSynopsis
from yw2oxmllib.docx.docx_export import DocxExport
from yw2oxmllib.docx.docx_characters import DocxCharacters
from yw2oxmllib.docx.docx_items import DocxItems
from yw2oxmllib.docx.docx_locations import DocxLocations


class Yw2msoExporter(YwCnvFf):
    """A converter for universal export from a yWriter 7 project.

    Public methods:
        export_from_yw(sourceFile, targetFile) -- Convert from yWriter project to other file format.

    Instantiate a Yw7File object as sourceFile and a
    Novel subclass object as targetFile for file conversion.
    Shows the 'Open' button after conversion from yw.

    Overrides the superclass constants EXPORT_SOURCE_CLASSES, EXPORT_TARGET_CLASSES.    
    """
    EXPORT_SOURCE_CLASSES = [Yw7File]
    EXPORT_TARGET_CLASSES = [DocxExport,
                             DocxBriefSynopsis,
                             DocxSceneDesc,
                             DocxChapterDesc,
                             DocxPartDesc,
                             DocxCharacters,
                             DocxLocations,
                             DocxItems,
                             # XlsxSceneList,
                             # XlsxCharList,
                             # XlsxLocList,
                             # XlsxItemList,
                             ]

    def export_from_yw(self, source, target):
        """Convert from yWriter project to other file format.

        Positional arguments:
            source -- YwFile subclass instance.
            target -- Any Novel subclass instance.
        
        Extends the super class method, showing an 'open' button after conversion.
        """
        super().export_from_yw(source, target)
        if self.newFile:
            self.ui.show_open_button()
        else:
            self.ui.hide_open_button()

