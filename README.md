# yw2oxml - yWriter to Open XML converter

This project was originally intended as an Open XML variant of the [yWriter to OpenOffice/LibreOffice standalone converter](https://peter88213.github.io/yW2OO/). 

For more information, see the [project homepage](https://peter88213.github.io/yw2oxml) with description and download instructions.

## Important

Please note that the program has not yet been extensively tested. To me, it's actually just a proof of concept. I probably won't develop the program further. Feel free to copy the project and modify it to your own liking.

### DOCX text document export

The yw2oxml script creates *docx* files, formatted in a "standard manuscript pages" layout. 

### XSLX spreadsheet export (not fully implemented)

In principle, it is also possible to export spreadsheets, such as scene lists or character lists. For this purpose there are some modules in the subdirectory `src\yw2oxmllib\xlsx`, which create XLSX table files according to the name, but still contain the code for OpenDocument ODS inside. 

To create proper *xlsx* documents, the following classes must be adapted first:

- OxmlFile
- XlsxFile
- XslxCharList
- XslXItemList
- XslxLocList
- XslxSceneList


### Conventions

See https://github.com/peter88213/PyWriter/blob/main/docs/conventions.md

Exceptions:
- No localization is required.
- The directory structure is modified to minimize dependencies.


### Development tools

- [Python](https://python.org) version 3.9
- [Eclipse IDE](https://eclipse.org) with [PyDev](https://pydev.org) and [EGit](https://www.eclipse.org/egit/)
- Apache Ant for building the application script

## License

yW2oxml is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
