#    Copyright 2018 Roel Zinkstok
#
#    This file is part of zink.
#
#    zink is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    zink is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with zink.  If not, see <http://www.gnu.org/licenses/>.


from .docxFrame import *
from .UnicodeToLaTeXLibrary import *


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

