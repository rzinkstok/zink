from docxFrame import *
from UnicodeToLaTeXLibrary import *


def processP(p):
    fec = FontEncodingConvertor()
    greektext = ''
    for node in p:
        if node.tag == W+'r':
            if node.hasProperty('rFonts', {W+"ascii": "Greek", W+"hAnsi": "Greek"}):
                greektext += node.getText()
    if len(greektext) > 0:
        print(greektext)
        print()
        for c in greektext:
            print(c, " -> ", ord(c), "%x" % ord(c), fec.getUnicode('greek', c))
        print()
        print()

