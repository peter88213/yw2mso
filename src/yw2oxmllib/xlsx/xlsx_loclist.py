"""Provide a class for XLSX location list export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from yw2oxmllib.xlsx.xlsx_file import XlsxFile


class XlsxLocList(XlsxFile):
    """XLSX location list representation."""
    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'

    _fileHeader = f'''{XlsxFile._DOCUMENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:number-columns-repeated="1014" table:default-cell-style-name="Default"/>
     <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="heading" office:value-type="string">
      <text:p>ID</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="heading" office:value-type="string">
      <text:p>Name</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="heading" office:value-type="string">
      <text:p>Description</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="heading" office:value-type="string">
      <text:p>Aka</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="heading" office:value-type="string">
      <text:p>Tags</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="heading" table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    _locationTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell office:value-type="string">
      <text:p>LcID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$AKA</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell table:number-columns-repeated="1014"/>
    </table:table-row>

'''
    _fileFooter = XlsxFile._DOCUMENT_XML_FOOTER 
