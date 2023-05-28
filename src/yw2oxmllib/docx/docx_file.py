"""Provide a generic class for DOCX file export.

Other DOCX file representations inherit from this class.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2oxml
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
import os
from string import Template
import locale

from pywriter.pywriter_globals import ERROR
from pywriter.file.file_export import FileExport
from yw2oxmllib.oxml.oxml_file import OxmlFile


class DocxFile(OxmlFile):
    """Generic Open XML text document representation."""

    EXTENSION = '.docx'
    # overwrites Novel.EXTENSION

    _OXML_COMPONENTS = ['[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml', 'word/_rels/document.xml.rels', 'word/styles.xml', 'word/document.xml', 'word/fontTable.xml', 'word/footer1.xml', 'word/settings.xml', ]

    _CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
    <Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
'''
    _APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
</Properties>
'''
    _DOCUMENT_XML_RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>
'''
    _STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">
    <w:docDefaults>
        <w:rPrDefault><w:rPr>
            <w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:eastAsia="Segoe UI" w:cs="Segoe UI"/>
            <w:color w:val="000000"/>
            <w:szCs w:val="2"/>
            <w:lang w:val="$Language-$Country" w:eastAsia="zxx" w:bidi="zxx"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <w:suppressAutoHyphens w:val="true"/>
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>
<w:qFormat/>
<w:pPr>
<w:widowControl w:val="false"/>
<w:suppressAutoHyphens w:val="true"/>
<w:overflowPunct w:val="false"/>
<w:bidi w:val="0"/>
<w:spacing w:lineRule="exact" w:line="414" w:before="0" w:after="0"/>
<w:jc w:val="left"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="Segoe UI" w:cs="Segoe UI"/>
<w:b w:val="false"/>
<w:color w:val="000000"/>
<w:kern w:val="0"/>
<w:sz w:val="24"/>
<w:szCs w:val="2"/>
<w:lang w:val="$Language-$Country" w:eastAsia="zxx" w:bidi="zxx"/>
</w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:spacing w:before="828" w:after="414"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:caps/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:spacing w:before="828" w:after="414"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:spacing w:before="414" w:after="414"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:spacing w:before="414" w:after="414"/></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading7"><w:name w:val="heading 7"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading8"><w:name w:val="heading 8"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading9"><w:name w:val="heading 9"/><w:basedOn w:val="Heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/><w:basedOn w:val="Normal"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:keepNext w:val="true"/><w:tabs><w:tab w:val="clear" w:pos="709"/></w:tabs><w:spacing w:lineRule="exact" w:line="414"/><w:jc w:val="center"/></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextFirstIndent"/><w:pPr><w:tabs><w:tab w:val="clear" w:pos="709"/></w:tabs></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="BodyTextFirstIndent"><w:name w:val="Body Text First Indent"/><w:basedOn w:val="BodyText"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="0" w:right="0" w:firstLine="283"/></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="BodyTextIndent"><w:name w:val="Body Text Indent"/><w:basedOn w:val="BodyText"/><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="283" w:right="0" w:hanging="0"/></w:pPr><w:rPr></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="heading10"><w:name w:val="heading 10"/><w:basedOn w:val="heading"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:outlineLvl w:val="8"/></w:pPr><w:rPr><w:b/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="Footer"/><w:basedOn w:val="Normal"/><w:pPr><w:suppressLineNumbers/><w:tabs><w:tab w:val="clear" w:pos="709"/><w:tab w:val="center" w:pos="4819" w:leader="none"/><w:tab w:val="right" w:pos="9639" w:leader="none"/></w:tabs><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Subtitle"/><w:qFormat/><w:pPr><w:suppressLineNumbers/><w:tabs><w:tab w:val="clear" w:pos="709"/></w:tabs><w:spacing w:lineRule="auto" w:line="480" w:before="0" w:after="0"/><w:ind w:left="0" w:right="0" w:hanging="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b w:val="false"/><w:caps/><w:kern w:val="0"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Title"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:rPr><w:b w:val="false"/><w:i/><w:caps w:val="false"/><w:smallCaps w:val="false"/><w:spacing w:val="0"/></w:rPr></w:style>
</w:styles>
'''
    _DOCUMENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14">
<w:body>
'''

    _DOCUMENT_XML_FOOTER = '''<w:sectPr w:rsidR="005C14FC" w:rsidSect="005C14FC">
    <w:footerReference w:type="default" r:id="rId7"/>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1814" w:right="1701" w:bottom="2380" w:left="1531" w:header="720" w:footer="1417" w:gutter="0"/>
    <w:cols w:space="0"/>
    </w:sectPr>
    </w:body>
    </w:document>
'''

    _FONT_TABLE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="Symbol"><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="Arial"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="Segoe UI"><w:charset w:val="01"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font>
<w:font w:name="Courier New"><w:charset w:val="01"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font>
<w:font w:name="Consolas"><w:charset w:val="01"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font></w:fonts>
'''
    _FOOTER1_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
    <w:p w:rsidR="005C14FC" w:rsidRDefault="005C14FC">
        <w:pPr>
            <w:pStyle w:val="Footer"/>
        </w:pPr>
        <w:fldSimple w:instr=" PAGE ">
            <w:r w:rsidR="00AA5003">
                <w:rPr><w:noProof/></w:rPr>
                <w:t>10</w:t>
            </w:r>
        </w:fldSimple>
    </w:p>
</w:ftr>
'''

    _SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
    <w:defaultTabStop w:val="709"/>
    <w:autoHyphenation w:val="true"/>
    <w:compat>
        <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    </w:compat>
    <w:rsids>
    <w:rsidRoot w:val="005C14FC"/>
        <w:rsid w:val="00331309"/>
        <w:rsid w:val="005C14FC"/>
        <w:rsid w:val="00AA5003"/>
    </w:rsids>
    <w:themeFontLang w:val="" w:eastAsia="" w:bidi=""/>
</w:settings>
'''

    def _set_up(self):
        """Helper method for ZIP file generation.

        Build the temporary directory containing the internal structure of an OXML file.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """

        # Generate the common OXML components.
        message = super()._set_up()
        if message.startswith(ERROR):
            return message

        os.mkdir(f'{self._tempDir}/word')
        os.mkdir(f'{self._tempDir}/word/_rels')

        #--- Generate docProps/app.xml.
        appMapping = dict(
        )
        template = Template(self._APP_XML)
        text = template.safe_substitute(appMapping)
        try:
            with open(f'{self._tempDir}/docProps/app.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Cannot write "app.xml"'

        #--- Generate word/document.xml.rels.
        try:
            with open(f'{self._tempDir}/word/_rels/document.xml.rels', 'w', encoding='utf-8') as f:
                f.write(self._DOCUMENT_XML_RELS)
        except:
            return f'{ERROR}Cannot write "document.xml.rels"'

        #--- Generate word/styles.xml.
        #  Set system language set as default language.
        lng, ctr = locale.getlocale()[0].split('_')
        stylesMapping = dict(
            Language=lng,
            Country=ctr,
       )
        template = Template(self._STYLES_XML)
        text = template.safe_substitute(stylesMapping)
        try:
            with open(f'{self._tempDir}/word/styles.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Cannot write "styles.xml"'

        #--- Generate word/fontTable.xml.
        try:
            with open(f'{self._tempDir}/word/fontTable.xml', 'w', encoding='utf-8') as f:
                f.write(self._FONT_TABLE_XML)
        except:
            return f'{ERROR}Cannot write "fontTable.xml"'

        #--- Generate word/footer1.xml.
        try:
            with open(f'{self._tempDir}/word/footer1.xml', 'w', encoding='utf-8') as f:
                f.write(self._FOOTER1_XML)
        except:
            return f'{ERROR}Cannot write "footer1.xml"'

        #--- Generate word/settings.xml.
        try:
            with open(f'{self._tempDir}/word/settings.xml', 'w', encoding='utf-8') as f:
                f.write(self._SETTINGS_XML)
        except:
            return f'{ERROR}Cannot write "settings.xml"'

        #--- Generate word/document.xml.
        self._originalPath = self._filePath
        self._filePath = f'{self._tempDir}/word/document.xml'
        message = FileExport.write(self)
        self._filePath = self._originalPath
        if message.startswith(ERROR):
            return message

        return 'DOCX structure generated.'

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        Overrides the superclass method.
        """
        if quick:
            # Just clean up a one-liner without sophisticated formatting.
            try:
                return text.replace('&', '&amp;').replace('>', '&gt;').replace('<', '&lt;')

            except AttributeError:
                return ''

        if text:
            text = self._remove_inline_code(text)

            # Remove comments.
            text = re.sub('\/\*.+?\*\/', '', text)

            # process italics and bold markup reaching across linebreaks
            italics = False
            bold = False
            newlines = []
            lines = text.split('\n')
            for line in lines:
                if italics:
                    line = f'[i]{line}'
                    italics = False
                while line.count('[i]') > line.count('[/i]'):
                    line = f'{line}[/i]'
                    italics = True
                while line.count('[/i]') > line.count('[i]'):
                    line = f'[i]{line}'
                line = line.replace('[i][/i]', '')
                if bold:
                    line = f'[b]{line}'
                    bold = False
                while line.count('[b]') > line.count('[/b]'):
                    line = f'{line}[/b]'
                    bold = True
                while line.count('[/b]') > line.count('[b]'):
                    line = f'[b]{line}'
                line = line.replace('[b][/b]', '')
                newlines.append(line)
            text = '\n'.join(newlines).rstrip()

            # Apply docx formatting.
            DOCX_REPLACEMENTS = [
                ('&', '&amp;'),
                ('>', '&gt;'),
                ('<', '&lt;'),
                ('\n\n', (2 * '</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t xml:space="preserve">')),
                ('\n', '</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="BodyTextFirstIndent"/></w:pPr><w:r><w:t xml:space="preserve">'),
                ('\r', '\n'),
                ('[i]', '</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">'),
                ('[/i]', '</w:t></w:r><w:r><w:t xml:space="preserve">'),
                ('[b]', '</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">'),
                ('[/b]', '</w:t></w:r><w:r><w:t xml:space="preserve">'),
            ]
            for yw, oxml in DOCX_REPLACEMENTS:
                text = text.replace(yw, oxml)

            # Remove highlighting, alignment,
            # strikethrough, and underline tags.
            text = re.sub('\[\/*[h|c|r|s|u]\d*\]', '', text)
        else:
            text = ''
        return text
