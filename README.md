# zink: docx2xelatex

A simple to use tool to convert Word docx files to XeLaTeX code, ready to compile.

### Dependencies

docx2xelatex uses python3, and uses lxml for all xml juggling. Lxml can be installed using pip:
```
$ pip install lxml
``` 

For compiling the resulting XeLaTeX file, of course you will need TeXLive, which contains TeX, LaTeX and XeLaTeX. 
TexLive can be downloaded for your platform from https://www.tug.org/texlive/

### Usage

The Word document should have been tagged with the appropriate styles:
- Headings
- Captions
- Table notes

```
$ cd <path of the zink source>
$ python runner.py <path of docx file>
```

### Supported objects

The following objects are supported:
- chapters/sections/subsections
- formatting (bold, italic, underline, subscript, superscript)
- footnotes
- figures
- tables
- captions for figures and tables (if tagged with the appropriate styles)
- notes for tables
- equations
- fields (fields must be updated in the document before conversion)
- strange scripts (e.g. greek)
- Unicode characters

### Why zink?

Together with my brother I used to have a company for typesetting PhD theses and the like.
We used this python code to convert from Word to XeLaTeX, and would then use our custom
XeLaTeX style files to create beautiful books. The name of our company was Zink Typografie,
which lend its name to our software as well. See http://www.zinktypografie.nl/ for more information.

