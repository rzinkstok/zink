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


import os
import zipfile
from lxml import etree as ET


class DocxFileError(Exception):
    pass


class DocxXMLError(Exception):
    pass


### DocX file class

class DocX:
    def __init__(self, path):
        self.checkFilename(path)
        self.path = path
        self.folder, self.filename = os.path.split(path)
        self.openFile()

    def checkFilename(self, path):
        """
        Check for empty or non-existing filename and filetype (by extension).
        """

        if path is None or not os.path.exists(path):
            raise DocxFileError("Invalid or empty file name "+path)
        ext = os.path.splitext(path)[1]
        if ext not in ['.docx', '.docm']:
            raise DocxFileError("Invalid file type "+ext)

    def openFile(self):
        """
        Opens the zipfile of a docx file and extracts the different xml files.
        """

        # Open the file as zipfile and extract the document XML file
        zf = zipfile.ZipFile(self.path)
        self.documentxml = DocumentXML(zf.read('word/document.xml'))
        self.docpropscorexml = DocPropsCoreXML(zf.read('docProps/core.xml'))

        # All other xml files are optional
        try:
            self.footnotesxml = FootnotesXML(zf.read('word/footnotes.xml'))
        except KeyError:
            self.footnotesxml = None
        try:
            self.relationsxml = RelationsXML(zf.read('word/_rels/document.xml.rels'))
        except KeyError:
            self.relationsxml = None
        try:
            self.numberingxml = NumberingXML(zf.read('word/numbering.xml'))
        except KeyError:
            self.numberingxml = None
        zf.close()

    def closeFile(self):
        pass

    def getDocumentInfo(self):
        pass

    def getFilename(self):
        return self.filename

    def getDocumentXML(self):
        return self.documentxml

    def getFootnotesXML(self):
        return self.footnotesxml

    def getRelationsXML(self):
        return self.relationsxml

    def getNumberingXML(self):
        return self.numberingxml

    def extractMediaFile(self, name, target):
        """
        Extract a file from the zipfile of the docx file to a specified target location.
        """

        zf = zipfile.ZipFile(self.path)
        zf.extract(name, target)
        zf.close()

    def getCreator(self):
        return self.docpropscorexml.getCreator()

    # Perhaps not needed??
    def analyzeStyles(self):
        """
        Returns list of all paragraph and run styles used in the document.
        """

        return self.documentxml.analyzeStyles()

    def getListType(self, numId, iLvl):
        return self.numberingxml.getListType(numId, iLvl)

    def getDocumentBody(self):
        return self.documentxml.getBody()


### XML file classes

class BaseDocxXML:
    """
    A general base class for docx xml documents.
    """

    def __init__(self, xmlstring):
        self.nsprefix = 'w'
        self.xml = ET.XML(xmlstring, OOXMLParser)

    def countNodes(self):
        """
        Returns the total node count of the xml document.
        """

        n = 0
        for node in self.xml.iter():
            n = n + 1
        return n

    def save(self, filename='', encoding='UTF-8', pretty_print=False):
        """
        Save the xml document to file with the specified encoding, possibly in a nice readable format.
        """

        if filename == '':
            filename = self.filename
        xmltree = ET.ElementTree(self.xml)
        xmltree.write(filename, encoding=encoding, standalone=True, pretty_print=pretty_print)

    def findAll(self, tag, addPrefix=True):
        """
        Find all elements in the document with the given tag.
        """

        return self.xml.findAll(tag, addPrefix)

    def findFirst(self, tag):
        """
        Find the first element in the document with the given tag.
        """

        return self.xml.findFirst(tag)

    def iter(self):
        """
        Iterate over all elements in the document.
        """

        return self.xml.iter()

class DocxXML(BaseDocxXML):
    """
    The base class for most xml documents found in a docx document.
    """

    def getParagraphs(self):
        return self.findAll('p')

    def analyzeStyles(self):
        """
        Return a list of all run and paragraph styles used in the document.
        """

        pstyles = []
        rstyles = []

        pss = self.findAll('pStyle')
        for ps in pss:
            stylename = ps.get('val')
            if stylename not in pstyles:
                pstyles.append(stylename)

        rss = self.findAll('rStyle')
        for rs in rss:
            stylename = rs.get('val')
            if stylename not in rstyles:
                rstyles.append(stylename)

        return {'runstyles':rstyles, 'paragraphstyles':pstyles}


class DocPropsCoreXML(DocxXML):
    def __init__(self, xmlstring=None, dir='docProps'):
        DocxXML.__init__(self, xmlstring)
        self.nsprefix = 'cp'
        self.dir = dir
        self.filename = os.path.join(self.dir, 'core.xml')

    def getCreator(self):
        cp = self.xml.findFirst('coreProperties')
        c = cp.findFirst('dc:creator')
        return c.text


class DocumentXML(DocxXML):
    """
    The class for holding the document.xml file from a docx file.
    """

    def __init__(self, xmlstring=None, dir='word'):
        DocxXML.__init__(self, xmlstring)
        self.dir = dir
        self.filename = os.path.join(self.dir, 'document.xml')

    def getBody(self):
        return self.findFirst('body')

    def getHeadings(self, styles=[]):
        headingdict = {}
        headinglist = []

        par = self.findFirst('p')
        while par is not None:
            parstyle = par.getStyle()
            if parstyle in styles:
                    headingdict[par] = parstyle
                    headinglist.append({'level': styles.index(parstyle), 'text': par.getText().strip()})
            par = par.getNextParagraph()

        print()
        for h in headinglist:
            for i in range(h['level']):
                print("\t", end='')
            print(h['text'])
        return headingdict


class FootnotesXML(DocxXML):
    """
    The class for holding the footnotes.xml file from a docx file.
    """

    def __init__(self, xmlstring=None, dir='word'):
        DocxXML.__init__(self, xmlstring)
        self.dir = dir
        self.filename = os.path.join(self.dir, 'footnotes.xml')

    def getFootnotes(self):
        return self.findFirst('footnotes')

    def getFootnote(self, fid):
        fns = self.findAll('footnote')
        for fn in fns:
            if fn.get('id') == fid:
                return fn
        return None


class RelationsXML(BaseDocxXML):
    """
    The class for holding the document.xml.rels file from a docx file.
    """

    def __init__(self, xmlstring=None, dir='word/_rels'):
        BaseDocxXML.__init__(self, xmlstring)
        self.nsprefix = 'rel'
        self.dir = dir
        self.filename = os.path.join(self.dir, 'document.xml.rels')

    def resolveRelation(self, rid):
        print("Resolving relation id "+str(rid))
        rels = self.findAll('Relationship')
        for rel in rels:
            if rel.get('Id') == rid:
                return rel.get('Target')
        return None


class NumberingXML(BaseDocxXML):
    """
    The class for holding the numbering.xml file from a docx file.
    """

    def __init__(self, xmlstring=None, dir='word'):
        BaseDocxXML.__init__(self, xmlstring)
        self.dir = dir
        self.filename = os.path.join(self.dir, 'numbering.xml')

    def resolveAbstractNumberingDefinition(self, anumid):
        anums = self.findAll('abstractNum')
        for anum in anums:
            if anum.get('abstractNumId') == anumid:
                return anum
        return None

    def resolveNumberingDefinition(self, numid):
        nums = self.findAll('num')
        for num in nums:
            if num.get('numId') == numid:
                return self.resolveAbstractNumberingDefinition(num.findFirst('abstractNumId').get('val'))
        return None

    def getListType(self, numid, ilvl):
        anumDef = self.resolveNumberingDefinition(numid)
        if anumDef is not None:
            lvls = anumDef.findAll('lvl')
            if len(lvls)>0:
                for lvl in lvls:
                    if ilvl == lvl.get('ilvl'):
                        return lvl.findFirst('numFmt').get('val')
        print("No list type found for numId={:s}, ilvl={:s}".format(numid, ilvl))
        return None


### Office Open XML base element class

class OfficeOpenXMLElement(ET.ElementBase):
    """
    The base element class for all elements in the Office Open XML standard
    """

    def _init(self):
        print("OOXML Element!")
        self.nsprefix = 'w'

    def buildNamespaceTag(self, tag):
        """
        Convert tags with prefixes to tags with namespaces; if no prefix is given, the nsprefix of current element is used.
        """

        tag = tag.split(':')
        if len(tag) == 1:
            prefix = self.nsprefix
            tag = tag[0]
        else:
            prefix = tag[0]
            tag = tag[1]

        tag = '{' + NSMAP[prefix] + '}' + tag
        return tag

    def buildPrefixTag(self, tag):
        """
        Ensures the tag includes a namespace prefix; if it does not have one, the prefix of the current element is used.
        """

        tag = tag.split(':')
        if len(tag) == 1:
            tag.insert(0, self.nsprefix)
        tag = ':'.join(tag)
        return tag

    def get(self, key, usePrefix=True, default=None):
        """
        Get an element attribute using the namespace prefix of the current element, or a user-specified prefix.
        """

        key = key.split(':')
        if len(key) == 1:
            prefix = self.nsprefix
        else:
            prefix = key[0]
            key.pop(0)

        # If the key contains :, it needs to be re-assembled
        key = ':'.join(key)

        if usePrefix:
            key = "{"+NSMAP[prefix]+"}"+key

        return ET.ElementBase.get(self, key, default)

    def doXpath(self, exp):
        """
        Runs the specified query on an element tree created using this element as root.
        """

        return ET.ElementTree(self).xpath(exp, namespaces=NSMAP)

    def findAll(self, tag, addPrefix=True):
        """
        Finds all elements in all descendants that have the specified (prefix) tag.
        """

        if addPrefix:
            tag = self.buildPrefixTag(tag)

        return self.doXpath('//'+tag)

    def findFirst(self, tag, addPrefix=True):
        """
        Finds the first element in all descendants that has the specified tag.
        """

        result = self.findAll(tag, addPrefix)
        if result:
            return result[0]
        else:
            return None

    def findAllInChildren(self, tag):
        """
        Finds all elements in the first generation children that have the specified (prefix) tag.
        """

        tag = self.buildNamespaceTag(tag)

        result = []
        cs = self.getchildren()
        for c in cs:
            if c.tag == tag:
                result.append(c)
        return result

    def findFirstInChildren(self, tag):
        """
        Finds the first element in the first generation children that has the specified tag.
        """

        result = self.findAllInChildren(tag)
        if result:
            return result[0]
        else:
            return None

    def drop_tree(self):
        """
        Remove the element from the xml tree with all its decendants.
        """

        parent = self.getparent()
        assert parent is not None
        parent.remove(self)

    def getTag(self):
        """
        Returns the element tag without the namespace.
        """

        ind = self.tag.find('}')
        if ind < 0:
            tag = self.tag
        else:
            tag = self.tag[ind+1:]

        return self.nsprefix+':'+tag


### The XML parser for OOXML documents

lookup = ET.ElementNamespaceClassLookup()
OOXMLParser = ET.XMLParser()
OOXMLParser.set_element_class_lookup(lookup)

NSMAP = {}