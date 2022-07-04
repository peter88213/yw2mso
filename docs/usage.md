[Project homepage](https://peter88213.github.io/yw2oxml)

------------------------------------------------------------------

## Command reference

-   [Export to docx](#export-to-docx)
-   [Brief synopsis](#brief-synopsis)
-   [Scene descriptions](#scene-descriptions)
-   [Chapter descriptions](#chapter-descriptions)
-   [Part descriptions](#part-descriptions)
-   [Character descriptions](#character-descriptions)
-   [Location descriptions](#location-descriptions)
-   [Item descriptions](#item-descriptions)


yWriter export to Open XML documents. 

# Instructions for use

## How to install yw2oxml

1. If you have already installed an older version of yw2oxml, please run the uninstaller for it. 

2. Unzip `yw2oxml_<version number>.zip` within your user profile.

3. Move into the `yw2oxml_<version number>` folder and run `setup.pyw` (double click).
   This will copy all needed files to the right places. 
   
The included installation script prompts you to create a shortcut on the desktop. 

You can either

- launch the program by double-clicking on the program/link icon, or
- launch the program by dragging a yWriter project file and dropping it on the program/link icon.

### Operation

#### Open a yWriter project

- If no yWriter project is specified by dragging and dropping on the program icon, the latest project selected is preset. You can change it with **File > Open** or **Ctrl-o**.

#### Close the ywriter project

- You can close the project without exiting the program with **File > Close**.
- If you open another project, the current project is automatically closed.

#### Exit 

- You can exit with **File > Exit** of **Ctrl-q**.


# Command reference

## Export to docx

This will load yWriter 7 chapters and scenes into a new Open XML
text document (docx).

-   The document is placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name>.docx`.
-   Text markup: Bold and italics are supported. Other highlighting such
    as underline and strikethrough are lost.
-   Only "normal" chapters and scenes are exported. Chapters and
    scenes marked "unused", "todo" or "notes" are not exported.
-   Only scenes that are intended for RTF export in yWriter are exported.
-   Comments in the text bracketed with slashes and asterisks (like
    `/* this is a comment */`) are removed.
-   Interspersed HTML, TEX, or RTF commands are removed.
-   Gobal variables and project variables are not resolved.
-   Chapter titles appear as first level heading if the chapter is
    marked as beginning of a new section in yWriter. Such headings are
    considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not
    marked as beginning of a new section. Such headings are considered
    as "chapter" headings.
-   Usually, scenes are separated by three asterisks. The first line is not
    indented.
-   Starting from the second paragraph, paragraphs begin with
    indentation of the first line.
-   Scenes marked "attach to previous scene" in yWriter appear like
    continuous paragraphs.



[Top of page](#top)

------------------------------------------------------------------------

## Brief synopsis

This will load a brief synopsis with chapter and scenes titles into a new
 Open XML text document (docx) placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name_brf_synopsis>.docx`.
-   Only "normal" chapters and scenes are exported. Chapters and
    scenes marked "unused", "todo" or "notes" are not exported.
-   Only scenes that are intended for RTF export in yWriter will be
    exported.
-   Chapter titles appear as first level heading if the chapter is
    marked as beginning of a new section in yWriter. Such headings are
    considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not
    marked as beginning of a new section. Such headings are considered
    as "chapter" headings.
-   Scene titles appear as plain paragraphs.



[Top of page](#top)

------------------------------------------------------------------------

## Scene descriptions

This will generate a new Open XML text document (docx) containing a
**full synopsis** with chapter titles and scene descriptions that can be
edited and written back to yWriter format. File name suffix is
`_scenes`.



[Top of page](#top)

------------------------------------------------------------------------

## Chapter descriptions

This will generate a new Open XML text document (docx) containing a
**brief synopsis** with chapter titles and chapter descriptions that can
be edited and written back to yWriter format. File name suffix is
`_chapters`.

**Note:** Doesn't apply to chapters marked
`This chapter begins a new section` in yWriter.



[Top of page](#top)

------------------------------------------------------------------------

## Part descriptions

This will generate a new Open XML text document (docx) containing a
**very brief synopsis** with part titles and part descriptions that can
be edited and written back to yWriter format. File name suffix is
`_parts`.

**Note:** Applies only to chapters marked
`This chapter  begins a new section` in yWriter.



## Character descriptions

This will generate a new Open XML text document (docx) containing
character descriptions, bio, goals, and notes that can be edited in Office
Writer and written back to yWriter format. File name suffix is
`_characters`.



[Top of page](#top)

------------------------------------------------------------------------

## Location descriptions

This will generate a new Open XML text document (docx) containing
location descriptions that can be edited in Office Writer and written
back to yWriter format. File name suffix is `_locations`.



[Top of page](#top)

------------------------------------------------------------------------

## Item descriptions

This will generate a new Open XML text document (docx) containing
item descriptions that can be edited in Office Writer and written back
to yWriter format. File name suffix is `_items`.



[Top of page](#top)

------------------------------------------------------------------------


## Installation path

The setup script installs *yw2oxml.pyw* in the user profile. This is the installation path on Windows: 

`c:\Users\<user name>\.pywriter\yw2oxml`

