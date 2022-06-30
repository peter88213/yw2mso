# yw2mso -- yWriter to MS Office converter

This project is meant as a MS Office variant of the [yWriter to OpenOffice/LibreOffice standalone converter](https://peter88213.github.io/yW2OO/). 

For more information, see the [project homepage](https://peter88213.github.io/yw2mso) with description and download instructions.

## Important

This is a work in progress. Contributions are welcome.

The yw2mso script creates DOCX files, but formatting issues may still occur. 

To create proper XLSX documents, the following classes must be adapted first:

- OxmlFile
- XlsxFile
- XslxCharList
- XslXItemList
- XslxLocList
- XslxSceneList


### Conventions

- Minimum Python version is 3.6. 
- The Python **source code formatting** follows widely the [PEP 8](https://www.python.org/dev/peps/pep-0008/) style guide, except the maximum line length, which is 120 characters here.

### Development tools

- [Python](https://python.org) version 3.9
- [Eclipse IDE](https://eclipse.org) with [PyDev](https://pydev.org) and [EGit](https://www.eclipse.org/egit/)
- [Apache Ant](https://ant.apache.org/) for building the application script

## License

yW2mso is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
