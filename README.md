# zink: docx2xelatex

A simple to use tool to convert Word docx files to XeLaTeX code, ready to compile.

### Dependencies

docx2xelatex uses python3, and needs a full TexLive installation. Furthermore, 
it uses lxml for all xml juggling. Lxml can be installed using pip:
```
$ pip install lxml
``` 

TexLive can be downloaded for your platform from https://www.tug.org/texlive/

### Usage:

The word document should have been tagged with the appropriate styles:
- Headings
- Captions
- Table notes
```
$ python runner.py <path of docx file>
```

### Supported objects

The following objects are supported:
- chapters/sections/subsections
- formatting (bold, italic, underline, subscript, superscript)
- footnotes
- figures
- tables
- captions for figures and tables (if tagged with teh appropriate styles)
- notes for tables
- equations
- fields (fields must be updated in the document before conversion)