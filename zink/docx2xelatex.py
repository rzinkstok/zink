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


import subprocess
import shutil
from .docxElements import *
from .UnicodeToLaTeXLibrary import *


class ConversionConfiguration(object):
    def __init__(self):
        self.style2heading = {
            'Kop1': 'chapter',
            'Heading1': 'chapter',

            'Kop2': 'section',
            'Heading2': 'section',

            'Kop3': 'subsection',
            'Heading3': 'subsection',

            'Kop4': 'subsubsection',
            'Heading4': 'subsubsection',

            'Kop5': 'paragraph',
            'Heading5': 'paragraph',

            'Kop7': 'subparagraph',
            'Heading7': 'subparagraph',

            'Kop9': 'sssssssSection',
            'Heading9': 'sssssssSection',
        }
        self.style2environment = {
            'quotation': 'quotation',
            'quote': 'quotation',
            'quotation_original': 'quotation_original',
            'quotation_translation': 'quotation_translation',
        }
        self.styleproperty2command = {
            'b': 'textbf',
            'i': 'textit',
            'u': 'ul',
            'superscript': 'textsuperscript',
            'subscript': 'textsubscript',
            'nor': 'mathrm',
            'double-struck': 'mathbb'
        }
        self.floatnames = ['Figure', 'Figuur', 'Afbeelding', 'Tabel', 'Table', 'Note']
        self.captionstyles = {
            'Caption': 'caption',
            'Bijschrift': 'caption',
            'Figurenote': 'caption'
        }
        self.tablecaptionstyles = {'TableCaption': 'caption'}
        self.tablenotestyles = ['Tablenote', 'Tablenote0']
        self.mathfunctions = {	'sin': '\\sin',
                                'cos': '\\cos',
                                'arccos': '\\arccos',
                                'csc': '\\csc',
                                'exp': '\\exp',
                                'min': '\\min',
                                'sinh': '\\sinh',
                                'arcsin': '\\arcsin',
                                'cosh': '\\cosh',
                                'ln': '\\ln',
                                'arctan': '\\arctan',
                                'cot': '\\cot',
                                'lim': '\\lim',
                                'log': '\\log',
                                'sec': '\\sec',
                                'tan': '\\tan',
                                'arg': '\\arg',
                                'coth': '\\coth',
                                'dim': '\\dim',
                                'max': '\\max',
                                'sin': '\\sin',
                                'tanh': '\\tanh'}
        self.mathaccents = {773: "\\overline"}
        self.convertable_image_extensions = ['.gif']
        self.incompatible_image_extensions = ['.emf', '.xml']

        self.useDirectStyling = True
        self.useTableColors = True

        self.supportedScriptLanguages = {'el-GR': 'greek'}
        self.supportedFonts = {'Greek': 'greek', 'Symbol': None}

        self.inHeading = False
        self.inMath = False
        self.inMathEnvironment = False
        self.inFootnote = False
        self.inStrangeScript = False
        self.inFont = False
        self.inTable = False
        self.inCaption = False
        self.inTableNote = False
        self.suppressFieldContents = False

        self.unicodelib = UnicodeToLaTeXLibrary()
        self.fontconvertor = FontEncodingConvertor()

        self.currentChapter = 0
        self.currentTable = 0
        self.references = []
        self.bookmarks = {}
        self.labels = {}

        self.captions = {}
        self.sortedcaptionkeys = []
        self.currentcaptionp = None
        self.currentcaptionlabels = {}
        self.cancelcurrentcaption = False

        self.tablecaptions = {}
        self.sortedtablecaptionkeys = []
        self.currenttablecaptiontbl = None

        self.tablenotes = {}
        self.sortedtablenotekeys = []
        self.currenttablenotetbl = None
        self.currenttablenotelabels = {}

        self.colors = []

        self.xelatexcode = XeLaTeXCode()

    def newChapter(self):
        self.currentChapter += 1
        self.currentTable = 0
        self.xelatexcode.newChapter("C{:d}".format(self.currentChapter))  # XXX

    def startTable(self, suppress=False):
        self.inTable = True
        if not suppress:
            self.currentTable += 1
            self.xelatexcode.toTables()

    def endTable(self, suppress=False):
        self.inTable = False
        if not suppress:
            self.xelatexcode.endToTables()

    def buildCaption(self):
        if self.currentcaptionp is not None:
            caption = "\caption{"
            for l in self.currentcaptionlabels.values():
                caption += "\label{"+l+"}"
            caption += self.xelatexcode.getString()
            caption += "}"
            self.xelatexcode.reset()

            self.captions[self.currentcaptionp] = caption
            self.sortedcaptionkeys.append(self.currentcaptionp)
            self.currentcaptionlabels = {}

            print()
            print(caption)

    def buildTableCaption(self):
        if self.currenttablecaptiontbl is not None:
            tcaption = ""
            for l in self.currentcaptionlabels.values():
                tcaption += "\label{"+l+"}"
            tcaption += self.xelatexcode.getString()
            self.xelatexcode.reset()

            self.tablecaptions[self.currenttablecaptiontbl] = tcaption
            self.sortedtablecaptionkeys.append(self.currenttablecaptiontbl)
            self.currentcaptionlabels = {}

            print()
            print(tcaption)

    def buildTableNote(self):
        if self.currenttablenotetbl is not None:
            tablenote = "\\tnote[]{"
            tablenote += self.xelatexcode.getString()
            tablenote += "}%"
            self.xelatexcode.reset()

            if not self.currenttablenotetbl in self.tablenotes:
                self.tablenotes[self.currenttablenotetbl] = []
            self.tablenotes[self.currenttablenotetbl].append(tablenote)
            self.sortedtablenotekeys.append(self.currenttablenotetbl)
            self.currenttablenotelabels = {}

            print()
            print(tablenote)

    def addColor(self, c):
        self.colors.append(c)


class XeLaTeXCode(object):
    def __init__(self):
        self.reset()

    def reset(self):
        self.data = {"Tables": "", "Frontmatter": ""}
        self.chapters = ["Tables", "Frontmatter"]
        self.curchap = -1
        self.toString = False
        self.xelatexstring = ""
        self.colors = []

    def startToString(self):
        self.xelatexstring = ""
        self.toString = True

    def endToString(self):
        self.toString = False

    def getString(self):
        return self.xelatexstring

    def addColor(self, c):
        if c not in self.colors:
            self.colors.append(c)

    def getColorDefinitions(self):
        if len(self.colors) == 0:
            return ''

        cdstring = "\n\n% Color definitions\n"
        for c in self.colors:
            r = int(c[0:2],16)/255.0
            g = int(c[2:4],16)/255.0
            b = int(c[4:6],16)/255.0
            cname = "RGB"+c
            cdstring += "\\definecolor{"+cname+"}{rgb}{"+str(r)+","+str(g)+","+str(b)+"}\n"
        cdstring += "\n"
        return cdstring

    def toTables(self):
        self.curchap = 0

    def endToTables(self):
        self.curchap = -1

    def append(self, s):
        if self.toString:
            self.xelatexstring += s
        else:
            self.data[self.chapters[self.curchap]] += s

    def addNL(self):
        self.append('\n')

    def newChapter(self, n):
        self.data[n]=""
        self.chapters.append(n)

    def writeRootFile(self, fpath, fnames):
        root = ""
        root += "\\documentclass[10pt]{memoir}\n"
        root += "\\usepackage{graphicx}\n"
        root += "\\usepackage{multirow}\n"
        root += "\\usepackage[table]{xcolor}\n"
        root += "\\usepackage{color}\n"
        root += "\\usepackage{ctable}\n"
        root += "\\begin{document}\n"

        root += self.getColorDefinitions()

        for f in fnames:
            # get filename (no path) without extension
            fn = os.path.splitext(os.path.split(f)[-1])[0]
            root += "\\include{" + fn + "}\n"

        root += "\\end{document}\n"

        print("Writing root file {:s}".format(fpath))
        fp = open(fpath, "w")
        fp.write(root)
        fp.close()

    def toFile(self, folder, basename):
        fnames = []
        for c in self.chapters:
            if c == "Tables" and self.data[c] == "":
                continue
            fname = basename + "_" + c + ".tex"
            fnames.append(fname)
            fpath = os.path.join(folder, fname)
            print("Writing content file {:s}".format(fpath))
            fp = open(fpath, "w")
            fp.write(self.data[c])
            fp.close()
        self.writeRootFile(os.path.join(folder, basename + "_root.tex"), fnames)


class docx2xelatex(object):
    def __init__(self, docx, verbose=False):
        self.docx = docx
        self.verbose = verbose
        self.docxfilename = self.docx.filename
        self.basefilename = os.path.splitext(self.docxfilename)[0]
        self.folder = self.docx.folder
        self.outputfolder = os.path.join(self.folder, self.basefilename)
        if not os.path.exists(self.outputfolder):
            os.makedirs(self.outputfolder)

        self.config = ConversionConfiguration()
        self.xelatexcode = self.config.xelatexcode

        if self.verbose:
            print()
            print("Converting {:s} to xelatex".format(self.docxfilename))
            print("Base filename: {:s}".format(self.basefilename))
            print("Folder: {:s}".format(self.docx.folder))
            print()

        self.docx.documentxml.save(filename=os.path.join(self.outputfolder, 'document_orig.xml'), pretty_print=True)
        if self.docx.footnotesxml is not None:
            self.docx.footnotesxml.save(filename=os.path.join(self.outputfolder, 'footnotes_orig.xml'), pretty_print=True)

    def buildReferenceList(self):
        itexts = []
        if self.verbose:
            print()
            print("==============================")
            print("  Building list of references")
            print("==============================")
            print()
            print("Document body:")
        d_itexts = self.findFieldInstructions(self.docx.documentxml.getBody())
        itexts.extend(i for i in d_itexts if i not in itexts)
        if self.docx.footnotesxml is not None:
            if self.verbose:
                print("Document footnotes:")
            f_itexts = self.findFieldInstructions(self.docx.footnotesxml.getFootnotes())
            itexts.extend(i for i in f_itexts if i not in itexts)

        self.config.references = []
        for i in itexts:
            fieldcodes = i.split()
            if fieldcodes[0] == "REF" or fieldcodes[0] == "NOTEREF" or fieldcodes[0] == "PAGEREF" and fieldcodes[1].find("_Toc")<0:
                self.config.references.append(fieldcodes[1])
        if self.verbose:
            print()
            print("Found {:d} references".format(len(self.config.references)))
            print()
            input()

    def buildBookmarks(self):
        if self.verbose:
            print()
            print("===============================")
            print("  Building list of bookmarks")
            print("===============================")
            print()
        bms = self.docx.documentxml.getBody().findAll('bookmarkStart')
        if self.verbose:
            print()
            print("Found {:d} bookmarks".format(len(bms)))
            print()
        for bm in bms:
            BookmarkProcessor(self, bm)
        if self.verbose:
            print()
            print("Found {:d} referenced bookmarks".format(len(self.config.bookmarks)))
            print()
            input()

    def buildCaptions(self):
        if self.verbose:
            print()
            print("==============================")
            print("  Building list of captions")
            print("==============================")
            print()

        lastnode = None
        openbookmarks = {}

        for n in self.docx.getDocumentBody().iter():
            if n.tag == W+'p' and n.isCaption(self.config.captionstyles):
                if lastnode.tag != W+'drawing':
                    lastnode = None
                CaptionProcessor(self, n, openbookmarks, lastnode)
                lastnode = None
            elif n.tag == W+'bookmarkStart':
                openbookmarks[n.getBookmarkId()] = n
            elif n.tag == W+'bookmarkEnd':
                openbookmarks.pop(n.getBookmarkId())
            else:
                lastnode = n

        if self.config.currentcaptionp is not None:
            self.config.buildCaption()
            self.config.currentcaptionp = None
            self.config.currentcaptionlabels = {}

        if self.verbose:
            print()
            print("Found {:d} captions".format(len(self.config.captions)))
            print()
            input()

    def buildTableCaptions(self):
        if self.verbose:
            print()
            print("==============================")
            print("  Building list of table captions")
            print("==============================")
            print()

        openbookmarks = {}
        for n in self.docx.getDocumentBody().iter():
            if n.tag == W+'p' and n.isCaption(self.config.tablecaptionstyles):
                TableCaptionProcessor(self, n, openbookmarks)
            if n.tag == W+'bookmarkStart':
                openbookmarks[n.getBookmarkId()] = n
            if n.tag == W+'bookmarkEnd':
                openbookmarks.pop(n.getBookmarkId())

        if self.config.currenttablecaptiontbl is not None:
            self.config.buildTableCaption()
            self.config.currenttablecaptiontbl = None
            self.config.currentcaptionlabels = {}

        if self.verbose:
            print()
            print("Found {:d} table captions".format(len(self.config.tablecaptions)))
            print()
            input()

    def buildTableNotes(self):
        if self.verbose:
            print()
            print("==============================")
            print("  Building list of table notes")
            print("==============================")
            print()

        for n in self.docx.getDocumentBody().iter():
            if n.tag == W+'p' and n.isTableNote(self.config.tablenotestyles):
                if n.getprevious().tag == W+'tbl':
                    TableNoteProcessor(self, n, n.getprevious())
            elif n.tag == W+'tbl':
                pass
            else:
                pass

        if self.config.currenttablenotetbl is not None:
            self.config.buildTableNote()
            self.config.currenttablenotetbl = None
            self.config.currenttablenotelabels = {}

        if self.verbose:
            print()
            print("Found {:d} table notes".format(len(self.config.tablenotes)))
            print()
            input()

    def findFieldInstructions(self, node):
        itexts = []
        flds = node.findAll('fldChar')
        fflds = [f for f in flds if f.isBegin()]
        if self.verbose:
            print("Found {:d} complex fields".format(len(fflds)))
        for f in fflds:
            itexts.append(f.getInstructionText())

        sflds = node.findAll('fldSimple')
        if self.verbose:
            print("Found {:d} simple fields".format(len(sflds)))
        for f in sflds:
            itexts.append(f.get('instr'))

        return itexts

    def process(self):
        self.buildReferenceList()
        self.buildBookmarks()
        self.xelatexcode.reset()
        self.buildTableNotes()
        self.xelatexcode.reset()
        self.buildCaptions()
        self.xelatexcode.reset()
        self.buildTableCaptions()
        self.xelatexcode.reset()

        node = self.docx.documentxml.getBody()[0]
        while node is not None:
            if node.tag == W+'p':
                ParagraphProcessor(self, node)
            elif node.tag == W+'tbl':
                TableProcessor(self, node)
            elif node.tag == W+'bookmarkStart':
                BookmarkProcessor(self, node)
            elif node.tag == W+'bookmarkEnd':
                pass
            elif node.tag == W+'sdt':
                pass
            elif node.tag == W+'sectPr':
                pass
            else:
                print("!! Document node {:s} ignored".format(node.tag))
            node = node.getnext()
        print()

        self.checkCaptions()

    def checkCaptions(self):
        if len(self.config.captions) > 0:
            print("Leftover captions!")
            print(self.config.captions)
            print()
        if len(self.config.tablecaptions) > 0:
            print("Leftover table captions!")
            print(self.config.tablecaptions)
            print()


    def writeLaTeXFiles(self):
        self.xelatexcode.toFile(self.outputfolder, self.basefilename)


# Processors

class Processor(object):
    def __init__(self, parentprocessor, node, verbose=False):
        self.parentprocessor = parentprocessor
        self.config = self.parentprocessor.config
        self.node = node
        self.docx = self.parentprocessor.docx
        self.xelatexcode = self.parentprocessor.xelatexcode
        self.outputfolder = parentprocessor.outputfolder
        self.verbose = verbose
        self.process()

    def process(self):
        pass


class ParagraphProcessor(Processor):
    def startHeading(self):
        h = self.config.style2heading[self.node.getStyle()]
        self.config.inHeading = True
        if h == 'chapter':
            self.config.newChapter()
            if self.verbose:
                print()
                print("New chapter: {:s}".format(self.node.getText()))
                print()
        self.xelatexcode.addNL()
        self.xelatexcode.append('\\'+h+'{')

    def endHeading(self):
        self.xelatexcode.append('}')
        self.xelatexcode.addNL()
        self.config.inHeading = False
        if self.node in self.config.bookmarks:
            h = self.config.style2heading[self.node.getStyle()]
            self.xelatexcode.append("\\label{" + h + ":" + self.config.bookmarks[self.node]["name"] + "}")
            self.xelatexcode.addNL()

    def startList(self):
        numId, iLvl = self.node.getNumbering()
        listType = self.docx.getListType(numId, iLvl)
        self.xelatexcode.addNL()
        if listType == 'bullet':
            self.xelatexcode.append('\\begin{itemize}')
        elif listType == 'decimal':
            self.xelatexcode.append('\\begin{enumerate}')
        elif listType == 'lowerRoman':
            self.xelatexcode.append('\\begin{enumerate}[(i)]')
        elif listType == 'lowerLetter':
            self.xelatexcode.append('\\begin{enumerate}[(a)]')
        else:
            print("List type {:s} not recognized, using itemize".format(listType))
            self.xelatexcode.append('\\begin{itemize}')

    def endList(self):
        numId, iLvl = self.node.getNumbering()
        listType = self.docx.getListType(numId, iLvl)
        self.xelatexcode.addNL()
        if listType == 'bullet':
            self.xelatexcode.append('\\end{itemize}')
        elif listType == 'decimal' or listType == 'lowerRoman' or listType == 'lowerLetter':
            self.xelatexcode.append('\\end{enumerate}')
        else:
            self.xelatexcode.append('\\end{itemize}')
        self.xelatexcode.addNL()

    def startListItem(self):
        if self.node.isFirstListItem():
            self.startList()
        self.xelatexcode.addNL()
        self.xelatexcode.append('\\item ')

    def endListItem(self):
        if self.node.isLastListItem():
            self.endList()

    def startEnvironmentIfNeeded(self):
        if self.node.isFirstWithEnvironment():
            self.xelatexcode.append("\\begin{" + self.node.getEnvironment(self.config.style2environment) + "}")

    def endEnvironmentIfNeeded(self):
        if self.node.isLastWithEnvironment():
            self.xelatexcode.append("\\end{" + self.node.getEnvironment(self.config.style2environment) + "}")

    def process(self):
        # Check heading, list, caption, footnote
        if self.node.isHeading(self.config.style2heading):
            if self.node.hasText():
                self.startHeading()
        elif self.node.isListItem():
            self.startListItem()
        elif not self.config.inCaption and self.node.isCaption(self.config.captionstyles):
            if self.node in self.config.captions.keys():
                self.config.xelatexcode.append(self.config.captions[self.node])
                self.xelatexcode.addNL()
                del(self.config.captions[self.node])
            return
        elif not self.config.inCaption and self.node.isCaption(self.config.tablecaptionstyles):
            return

        if self.node.isEnvironment(self.config.style2environment):
            self.startEnvironmentIfNeeded()

        if len(self.node) > 0:
            node = self.node[0]
            while node is not None:
                if node.tag == W+'r':
                    RunProcessor(self, node)
                elif node.tag == W+'pPr':
                    pass
                elif node.tag == M+'oMathPara':
                    MathParagraphProcessor(self, node)
                elif node.tag == M+'oMath':
                    MathProcessor(self, node)
                elif node.tag == W+'bookmarkStart' or node.tag == W+'bookmarkEnd':
                    pass
                elif node.tag == W+'fldSimple':
                    FieldSimpleProcessor(self, node)
                elif node.tag == W+'smartTag':
                    SubProcessor(self, node)
                elif node.tag == W+'hyperlink':
                    SubProcessor(self, node)
                else:
                    print("!! Paragraph subnode {:s} ignored".format(node.tag))
                node = node.getnext()

        if self.node.isEnvironment(self.config.style2environment):
            self.endEnvironmentIfNeeded()

        if self.node.isHeading(self.config.style2heading):
            if self.node.hasText():
                self.endHeading()
        elif self.node.isListItem():
            self.endListItem()

        if not self.config.inFootnote and not self.config.inTable and not self.config.inCaption and not self.config.inTableNote:
            self.xelatexcode.addNL()
            self.xelatexcode.addNL()


class RunProcessor(Processor):
    def startStyleIfNeeded(self, s):
        if self.node.isFirstWithStyleProperty(s):
            self.xelatexcode.append('\\' + self.config.styleproperty2command[s] + '{')

    def endStyleIfNeeded(self, s):
        if self.node.isLastWithStyleProperty(s):
            self.xelatexcode.append('}')

    def startStyles(self):
        if self.node.isBold(self.config.useDirectStyling):
            self.startStyleIfNeeded('b')
        if self.node.isItalic(self.config.useDirectStyling):
            self.startStyleIfNeeded('i')
        if self.node.isUnderline(self.config.useDirectStyling):
            self.startStyleIfNeeded('u')
        if self.node.isSuperScript():
            self.startStyleIfNeeded('superscript')
        if self.node.isSubScript():
            self.startStyleIfNeeded('subscript')

    def endStyles(self):
        if self.node.isBold(self.config.useDirectStyling):
            self.endStyleIfNeeded('b')
        if self.node.isItalic(self.config.useDirectStyling):
            self.endStyleIfNeeded('i')
        if self.node.isUnderline(self.config.useDirectStyling):
            self.endStyleIfNeeded('u')
        if self.node.isSuperScript():
            self.endStyleIfNeeded('superscript')
        if self.node.isSubScript():
            self.endStyleIfNeeded('subscript')

    def processLanguage(self):
        language = self.node.getLanguage()
        if (language is not None) and (language in self.config.supportedScriptLanguages.keys()) and (not self.config.inStrangeScript):
            self.startStrangeScript(language)
        if (language is None or language not in self.config.supportedScriptLanguages.keys()) and self.config.inStrangeScript:
            self.stopStrangeScript()

    def startStrangeScript(self, language):
        self.config.inStrangeScript = self.config.supportedScriptLanguages[language]
        self.xelatexcode.append("\\begin{" + self.config.inStrangeScript + "}")

    def stopStrangeScript(self):
        self.xelatexcode.append("\\end{" + self.config.inStrangeScript + "}")
        self.config.inStrangeScript = False

    def processFont(self):
        font = self.node.getFont()
        if font is not None and font in self.config.supportedFonts.keys() and not self.config.inFont:
            self.startFont(font)
        if (font is None or font not in self.config.supportedFonts.keys()) and self.config.inFont:
            self.stopFont(font)

    def startFont(self, font):
        self.config.inFont = font #self.config.supportedFonts[font]
        if font in self.config.supportedFonts.keys() and self.config.supportedFonts[font]:
            self.xelatexcode.append("\\begin{" + self.config.supportedFonts[font] + "}")

    def stopFont(self, font):
        if font in self.config.supportedFonts.keys() and self.config.supportedFonts[font]:
            self.xelatexcode.append("\\end{" + self.config.supportedFonts[font] + "}")
        self.config.inFont = False

    def process(self):
        self.processLanguage()
        self.processFont()
        self.startStyles()

        if len(self.node) > 0:
            node = self.node[0]
            while node is not None:
                if node.tag == W+'t':
                    TextProcessor(self, node)
                elif node.tag == W+'footnoteReference':
                    FootnoteProcessor(self, node)
                elif node.tag == W+'footnoteRef':
                    pass
                elif node.tag == M+'oMathPara':
                    MathParagraphProcessor(self, node)
                elif node.tag == W+'tab':
                    self.xelatexcode.append(" ")
                elif node.tag == W+'br':
                    self.xelatexcode.append(" ")
                elif node.tag == W+'softHyphen':
                    self.xelatexcode.append("\-") # Maybe not: this prevents hyphenation elsewhere in the word
                elif node.tag == W+'rPr' or node.tag == W+'lastRenderedPageBreak':
                    pass
                elif node.tag == W+'pict':
                    print("!! pict element not supported! Convert to pdf image.")
                elif node.tag == W+'drawing':
                    DrawingProcessor(self, node)
                elif node.tag == W+'fldChar':
                    FieldCharProcessor(self, node)
                elif node.tag == W+'instrText':
                    pass
                elif node.tag == W+'sym':
                    SymProcessor(self, node)
                elif node.tag == W+'noBreakHyphen':
                    self.xelatexcode.append("{-}")
                # bookmarkStart, bookmarkEnd, sym, lastRenderedPageBreak,
                elif node.tag == MC+'AlternateContent':
                    AlternateContentProcessor(self, node)
                else:
                    print("!! Run subnode {:s} ignored".format(node.tag))
                node = node.getnext()

        self.endStyles()


class TextProcessor(Processor):
    def process(self):
        LaTeXizer(self, self.node.getText())


class LaTeXizer(object):
    def __init__(self, parentprocessor, s, isString=False):
        self.parentprocessor = parentprocessor
        self.config = self.parentprocessor.config
        self.s = s
        self.isString = isString
        if self.isString:
            self.xelatexcode = ""
        else:
            self.xelatexcode = self.parentprocessor.xelatexcode
        self.process()

    def getXeLaTeXString(self):
        if self.isString:
            return self.xelatexcode
        else:
            return None

    def process(self):
        if self.config.suppressFieldContents:
            return

        for c in self.s:
            if self.config.inStrangeScript or self.config.inFont:
                self.addCode(self.XeLaTeXizeChar(c))
            elif self.config.inMath or self.config.inMathEnvironment:
                self.addCode(self.LaTeXizeMathChar(c))
            else:
                self.addCode(self.LaTeXizeChar(c))

    def addCode(self, s):
        if self.isString:
            self.xelatexcode += s
        else:
            self.xelatexcode.append(s)

    def XeLaTeXizeChar(self, c):
        # For strange scripts
        if self.config.inFont:
            c = self.config.fontconvertor.getUnicode(self.config.inFont, c)
        try:
            xc = self.config.unicodelib.getXeLaTeXChar(c)
        except UnicodeToLaTeXLibraryError:
            print("XeLaTeX character missing from {:s}".format(self.s))
            self.config.unicodelib.insertXeLaTeXChar(c)
            xc = self.config.unicodelib.getXeLaTeXChar(c)
        return xc

    def LaTeXizeChar(self, c):
        try:
            lc = self.config.unicodelib.getLaTeXChar(c)
        except UnicodeToLaTeXLibraryError:
            print("LaTeX character missing from {:s}".format(self.s))
            self.config.unicodelib.insertLaTeXChar(c)
            lc = self.config.unicodelib.getLaTeXChar(c)
        return lc

    def LaTeXizeMathChar(self, c):
        try:
            mc = self.config.unicodelib.getLaTeXMathChar(c)
        except UnicodeToLaTeXLibraryError:
            print("LaTeX math character missing from {:s}".format(self.s))
            self.config.unicodelib.insertLaTeXMathChar(c)
            mc = self.config.unicodelib.getLaTeXMathChar(c)
        return mc


class SymProcessor(Processor):
    def process(self):
        c = chr(self.s.getSymbolCode())
        f = self.s.getSymbolFont()
        uc = self.config.fontconvertor.getUnicode(f, c)
        l = LaTeXizer(self, uc, isString=True)
        s = l.getXeLaTeXString()
        if self.verbose:
            print("Sym: {:s}".format(s))
        self.xelatexcode.append(s)


class FootnoteProcessor(Processor):
    def process(self):
        fid = self.node.getFootnoteId()
        footnote = self.docx.footnotesxml.getFootnote(fid)

        self.config.inFootnote = True
        self.xelatexcode.append('\\footnote{')

        if fid in self.config.bookmarks:
            self.xelatexcode.append("\\label{footnote:"+self.config.bookmarks[fid]["name"]+"}")

        node = footnote[0]

        while node is not None:
            if node.tag == W+'p':
                ParagraphProcessor(self, node)
            elif node.tag == W+'tbl':
                TableProcessor(self, node)
            else:
                print("!! Node {:s} ignored in footnote id {:d}".format(node.tag, fid))
            node = node.getnext()

        self.xelatexcode.append('}')
        self.config.inFootnote = False


class TableProcessor(Processor):
    def arabicToAlphabetic(self, num):
        a, b = divmod(num, 26)
        if a > 0:
            res = chr(ord("@")+a)
        else:
            res = ""
        res += chr(ord("@")+b)
        if res == "@":
            return ""
        return res

    def process(self):
        suppressTable = not self.node.hasText('()0123456789', ['Caption', 'Figurenote', 'Tablenote', 'Tablenote0'])

        self.config.startTable(suppressTable)

        if not suppressTable:
            tableid = "tabc"+self.arabicToAlphabetic(self.config.currentChapter)+"t"+self.arabicToAlphabetic(self.config.currentTable)
            if self.node in self.config.tablenotes:
                tablenotes = self.config.tablenotes[self.node]
            else:
                tablenotes = []
            self.xelatexcode.append("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
            self.xelatexcode.addNL()
            self.xelatexcode.append("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
            self.xelatexcode.addNL()
            self.xelatexcode.addNL()
            self.xelatexcode.addNL()
            self.xelatexcode.append("\\newcommand{\\"+tableid+"}{%")
            self.xelatexcode.addNL()

            tw = self.node.getWidth()
            if tw is None or tw == 'auto':
                tablewidth = "1.0\\textwidth"
            elif tw[-1] == '%':
                tablewidth = "{:.1f}\\textwidth".format(0.01*float(tw[:-1]))
            else:
                tablewidth = "{:.1f}mm".format(float(tw))

            if self.config.inFootnote:
                self.xelatexcode.append("\\noindent{\\tiny\\begin{tabularx}{"+tablewidth+"}{%")
                self.xelatexcode.addNL()
            else:
                self.xelatexcode.append("\\ctable[%")
                self.xelatexcode.addNL()
                self.xelatexcode.append("\twidth="+tablewidth+",%")
                self.xelatexcode.addNL()
                #self.xelatexcode.append("\tdoinside=\\tablefont\\tiny,%")
                self.xelatexcode.append("\tdoinside=\\tiny,%")
                self.xelatexcode.addNL()
                self.xelatexcode.append("\tcaption={")
                if self.node in self.config.tablecaptions:
                    self.xelatexcode.append(self.config.tablecaptions[self.node])
                    del(self.config.tablecaptions[self.node])
                self.xelatexcode.append("}%")
                self.xelatexcode.addNL()
                self.xelatexcode.append("]%")
                self.xelatexcode.addNL()
                self.xelatexcode.append("{%")
                self.xelatexcode.addNL()

            ncols = self.node.getNumberOfLogicalColumns()
            nrows = self.node.getNumberOfRows()

            for i in range(ncols):
                self.xelatexcode.append("\t>{\\raggedright\\arraybackslash}X%")
                self.xelatexcode.addNL()
            self.xelatexcode.append("}%")
            self.xelatexcode.addNL()

            if not self.config.inFootnote:
                self.xelatexcode.append("{%")
                self.xelatexcode.addNL()
                for tn in tablenotes:
                    self.xelatexcode.append(tn)
                    self.xelatexcode.addNL()
                self.xelatexcode.append("}%")
                self.xelatexcode.addNL()
                self.xelatexcode.append("{%")
                self.xelatexcode.addNL()

            self.xelatexcode.append("\\toprule")
            self.xelatexcode.addNL()


        # Table body
        rows = self.node.getRows()
        rowcounter = 0
        for r in rows:
            # check for gridBefore
            if not suppressTable:
                for i in range(r.getGridBefore()):
                    self.xelatexcode.append(" & ")

            columncounter = 0
            cells = r.getCells()
            for c in cells:
                if columncounter > 0 and not suppressTable:
                    self.xelatexcode.append(" & ")

                # check for columnspans
                colspanning = c.getColumnSpanning()
                if (colspanning > 1) and not suppressTable:
                    self.xelatexcode.append('\\multicolumn{'+str(colspanning)+'}{c}{')

                # check for rowspans
                rowspanning = c.getRowSpanning()
                if (rowspanning > 1) and not suppressTable:
                    self.xelatexcode.append("\\multirow{"+str(rowspanning)+"}*{")

                # check cell color
                if self.config.useTableColors:
                    ccolor = c.getBackgroundColor()
                    if ccolor is not None:
                        self.xelatexcode.addColor(ccolor)
                        self.xelatexcode.append("\\cellcolor{RGB"+ccolor+"} ")

                node = c[0]
                while node is not None:
                    if node.tag == W+'p':
                        ParagraphProcessor(self, node)
                    elif node.tag == W+'tbl':
                        print("Table within table!")
                        TableProcessor(self, node)
                    elif node.tag == W+'tcPr':
                        #self.processTableCellProperties(pnode)
                        pass
                    else:
                        print("!! Table cell subnode {:s} ignored".format(node.tag))
                    node = node.getnext()

                # End row spanning
                if (rowspanning > 1) and not suppressTable:
                    self.xelatexcode.append("}")

                # End column spanning
                if (colspanning > 1) and not suppressTable:
                    self.xelatexcode.append("}")

                columncounter += 1

            if not suppressTable:
                self.xelatexcode.append("\\tabularnewline ")
                if rowcounter == 0 and nrows >1:
                    self.xelatexcode.append("\\midrule")
                self.xelatexcode.addNL()
            rowcounter += 1

        # End table
        if not suppressTable:
            self.xelatexcode.append("\\bottomrule")
            self.xelatexcode.addNL()

            if self.config.inFootnote:
                self.xelatexcode.append("\\end{tabularx}}}")
            else:
                self.xelatexcode.append("}}")

            self.xelatexcode.addNL()
            self.xelatexcode.addNL()

        self.config.endTable(suppressTable)

        if not suppressTable:
            self.xelatexcode.addNL()
            self.xelatexcode.append("\\"+tableid)
            self.xelatexcode.addNL()


class MathParagraphProcessor(Processor):
    def process(self):
        self.config.inMathEnvironment = True

        self.xelatexcode.addNL()
        self.xelatexcode.addNL()
        self.xelatexcode.append("\\begin{equation}")
        self.xelatexcode.addNL()

        if self.node in self.config.bookmarks:
            self.xelatexcode.append("\\label{equation:" + self.config.bookmarks[self.node]["name"] + "}")
            self.xelatexcode.addNL()

        node = self.node[0]
        while node is not None:
            if node.tag == M + 'oMathParaPr':
                # Process properties (only justification)
                pass
            if node.tag == M + 'oMath':
                MathProcessor(self, node)
            else:
                print("!! Math paragraph subnode {:s} ignored".format(node.tag))
            node = node.getnext()

        self.xelatexcode.addNL()
        self.xelatexcode.append("\\end{equation}")
        self.xelatexcode.addNL()

        self.config.inMathEnvironment = False


class MathProcessor(Processor):
    def process(self):
        if not self.config.inMathEnvironment:
            self.config.inMath = True

        mlc = self.processMathElement(self.node)

        if not self.config.inMathEnvironment:
            self.config.inMath = False

        if mlc.strip() != "":
            if not self.config.inMathEnvironment:
                self.xelatexcode.append("$")

            self.xelatexcode.append(mlc)

            if not self.config.inMathEnvironment:
                self.xelatexcode.append("$")

    def processMathElement(self, m):
        latexcode = ""
        if len(m)<1:
            return latexcode
        node = 	m[0]
        while node is not None:
            if node.tag == M + 'r':
                mrp = MathRunProcessor(self, node)
                latexcode += mrp.getXeLaTeXString()

            elif node.tag == M + 'sSub':
                e = self.processMathElement(node.findFirstInChildren('e'))
                sub = self.processMathElement(node.findFirstInChildren('sub'))

                latexcode += " " + e
                if sub != "":
                    latexcode += "_{" + sub + "} "

            elif node.tag == M+ 'sSup':
                e = self.processMathElement(node.findFirstInChildren('e'))
                sup = self.processMathElement(node.findFirstInChildren('sup'))

                latexcode += " " + e
                if sup.strip() != "":
                    while sup[0] == "'":
                        latexcode += sup[0]
                        if len(sup)>1:
                            sup = sup[1:]
                        else:
                            sup = ""
                            break
                    if sup.strip() != "":
                        latexcode += "^{" + sup + "}"

            elif node.tag == M + 'sSubSup':
                e = self.processMathElement(node.findFirstInChildren('e'))
                sup = self.processMathElement(node.findFirstInChildren('sup'))
                sub = self.processMathElement(node.findFirstInChildren('sub'))

                latexcode += " " + e
                if sup.strip() != "":
                    while sup[0] == "'":
                        latexcode += sup[0]
                        if len(sup)>1:
                            sup = sup[1:]
                        else:
                            sup = ""
                            break
                    if sup.strip() != "":
                        latexcode += '^{' + sup + '}'
                if sub != '':
                    latexcode += '_{' + sub + '} '

            elif node.tag == M + 'd':
                dPr = node.findFirstInChildren('dPr')

                begChr = LaTeXizer(self, "(", isString=True).getXeLaTeXString()
                sepChr = LaTeXizer(self, "|", isString=True).getXeLaTeXString()
                endChr = LaTeXizer(self, ")", isString=True).getXeLaTeXString()

                if dPr is not None:
                    bC = dPr.findFirstInChildren('begChr')
                    if bC is not None:
                        begChr = LaTeXizer(self, bC.get('val'), isString=True).getXeLaTeXString()
                    eC = dPr.findFirstInChildren('endChr')
                    if eC is not None:
                        endChr = LaTeXizer(self, eC.get('val'), isString=True).getXeLaTeXString()
                    sC = dPr.findFirstInChildren('sepChr')
                    if sC is not None:
                        sepChr = LaTeXizer(self, sC.get('val'), isString=True).getXeLaTeXString()

                latexcode += begChr #latexcode + ' \\left' + begChr + ' '


                subnodes = node.findAllInChildren('e')
                for subnode in subnodes:
                    latexcode += self.processMathElement(subnode)
                    if subnode is not subnodes[-1]:
                        latexcode += " " + sepChr + " "
                latexcode += endChr + " "

            elif node.tag == M + 'nary':
                naryPr = node.findFirstInChildren('naryPr')

                chr = "" #LaTeXizer(self, "^", isString=True).getXeLaTeXString()
                limLoc = 'undOvr'
                if naryPr is not None:
                    c = naryPr.findFirstInChildren('chr')
                    if c is not None:
                        chr = LaTeXizer(self, c.get('val'), isString=True).getXeLaTeXString()
                    ll = naryPr.findFirstInChildren('limLoc')
                    if ll is not None:
                        limLoc = ll.get('val')

                # use limLoc, \limits?

                e = self.processMathElement(node.findFirstInChildren('e'))
                sup = self.processMathElement(node.findFirstInChildren('sup'))
                sub = self.processMathElement(node.findFirstInChildren('sub'))

                latexcode += " " + chr
                if sub != "":
                    latexcode += "_{" + sub + "} "
                if sup != "":
                    latexcode += "^{" + sup + "} "
                latexcode += e + " "

            elif node.tag == M + 'f':
                fPr = node.findFirstInChildren('fPr')

                ftype = "bar"
                if fPr is not None:
                    t = fPr.findFirstInChildren('type')
                    if t is not None:
                        ftype = t.get('val')

                num = self.processMathElement(node.findFirstInChildren('num'))
                den = self.processMathElement(node.findFirstInChildren('den'))

                if ftype == "bar":
                    latexcode += " \\frac{" + num + "}{" + den + "} "
                elif ftype == "lin" or ftype == "skw":
                    latexcode += " " + num + "/" + den + " "
                elif ftype == 'noBar':
                    latexcode += " \\begin{matrix} " + num + " \\\\ " + den + " \\end{matrix} "

            elif node.tag == M + 'func':
                funcPr = node.findFirstInChildren('funcPr')

                fName = self.processMathElement(node.findFirstInChildren('fName'))
                e = self.processMathElement(node.findFirstInChildren('e'))

                if fName in self.config.mathfunctions:
                    fName = self.config.mathfunctions[fName]

                latexcode += " " + fName + " " + e + " "

            elif node.tag == M + 'acc':
                accPr = node.findFirstInChildren('accPr')

                chr = "^" #LaTeXizer(self, "^", isString=True).getXeLaTeXString()
                if accPr is not None:
                    c = accPr.findFirstInChildren('chr')
                    if c is not None:
                        chr = c.get('val')

                e = self.processMathElement(node.findFirstInChildren('e'))

                latexcode += self.addMathAccent(e, chr)

            elif node.tag == M + 'ctrlPr':
                pass

            else:
                print("!! Math element subnode {:s} ignored".format(node.tag))
            node = node.getnext()

        return latexcode

    def addMathAccent(self, e, chr):
        if ord(chr) in self.config.mathaccents:
            return self.config.mathaccents[ord(chr)] + "{" + e + '}'
        else:
            raise ValueError("Math accent character {:d} missing".format(ord(chr)))

class MathRunProcessor(Processor):
    def __init__(self, parentprocessor, r):
        self.parentprocessor = parentprocessor
        self.config = self.parentprocessor.config
        self.r = r
        self.xelatexstring = ""
        self.process()

    def getXeLaTeXString(self):
        return self.xelatexstring

    def process(self):
        self.startStyles()

        node = self.r[0]
        while node is not None:
            if node.tag == M + 't':
                mtp = MathTextProcessor(self, node)
                self.xelatexstring += mtp.getXeLaTeXString()
            elif node.tag == M + 'rPr' or node.tag == W + 'rPr':
                pass
            elif node.tag == W + 'br':
                pass
            else:
                print("!! Math run subnode {:s} ignored".format(node.tag))
            node = node.getnext()

        self.endStyles()

    def startStyles(self):
        if self.node.isRoman():
            self.startStyleIfNeeded('nor')
        if self.node.isDoubleStruck():
            self.startStyleIfNeeded('double-struck')

    def startStyleIfNeeded(self, s):
        if self.node.isFirstWithStyleProperty(s):
            self.xelatexstring += "\\" + self.config.styleproperty2command[s] + "{"

    def endStyles(self):
        if self.node.isRoman():
            self.endStyleIfNeeded('nor')
        if self.node.isDoubleStruck():
            self.endStyleIfNeeded('double-struck')

    def endStyleIfNeeded(self, s):
        if self.node.isLastWithStyleProperty(s):
            self.xelatexstring += "}"


class MathTextProcessor(Processor):
    def __init__(self, parentprocessor, t):
        self.parentprocessor = parentprocessor
        self.config = self.parentprocessor.config
        self.t = t
        self.xelatexstring = ""
        self.process()

    def getXeLaTeXString(self):
        return self.xelatexstring

    def process(self):
        l = LaTeXizer(self, self.t.getText(), isString=True)
        self.xelatexstring += l.getXeLaTeXString()


class DrawingProcessor(Processor):
    def process(self):
        commented = False
        latexname = None

        refID = self.node.getImageReferenceId()
        if refID is None:
            print("Skipping drawing (no reference ID found)")
            self.xelatexcode.addNL()
            self.xelatexcode.addNL()
        else:
            sx, sy = self.node.getImageSize()

            relxml = self.docx.getRelationsXML()
            zippath = os.path.join('word', relxml.resolveRelation(refID))

            filename = os.path.split(zippath)[-1]

            mediadirname = "media"
            mediapath = os.path.join(self.parentprocessor.outputfolder, "media")
            if not os.path.exists(mediapath):
                os.mkdir(mediapath)
            self.docx.extractMediaFile(zippath, mediapath)

            oldname = os.path.join(mediapath, zippath)
            newname = os.path.join(mediapath, filename)
            latexname = os.path.join(mediadirname, filename)
            os.rename(oldname, newname)
            shutil.rmtree(os.path.join(mediapath, 'word'))

            if os.path.splitext(filename)[1] in self.config.incompatible_image_extensions:
                commented = True
            elif os.path.splitext(filename)[1] in self.config.convertable_image_extensions:
                pdffilename = os.path.splitext(filename)[0]+".pdf"
                pdfpath = os.path.join(mediapath, pdffilename)
                subprocess.Popen(["/usr/bin/sips",  "-s",  "format",  "pdf", newname, "--out", pdfpath])
                latexname = os.path.join(mediadirname, pdffilename)

        self.xelatexcode.addNL()
        if commented:
            self.xelatexcode.append("%")
        self.xelatexcode.append("\\begin{")
        self.xelatexcode.append("figure")
        self.xelatexcode.append("}")
        self.xelatexcode.addNL()
        if commented:
            self.xelatexcode.append("%")
        if latexname is not None:
            self.xelatexcode.append("\\includegraphics[width={:.2f}cm]".format(sx)+"{"+latexname+"}")
        else:
            self.xelatexcode.append("%% Drawing omitted")
        self.xelatexcode.addNL()

        ### Caption
        if self.node in self.config.captions:
            self.xelatexcode.append("\\caption{")
            self.xelatexcode.append(self.config.captions[self.node])
            self.xelatexcode.append("}")
            self.xelatexcode.addNL()
            del(self.config.captions[self.node])

        if commented:
            self.xelatexcode.append("%")
        self.xelatexcode.append("\\end{")
        self.xelatexcode.append("figure")
        self.xelatexcode.append("}")
        self.xelatexcode.addNL()


class BookmarkProcessor(Processor):
    def process(self):
        bid = self.node.getBookmarkId()
        bname = self.node.getBookmarkName()
        if bname in self.config.references:
            if self.verbose:
                print("=========================")
                print("Bookmark id: "+bid+"; name: "+bname)
            btype, blink = self.node.getBookmarkType(self.config.style2heading.keys(), self.config.floatnames, self.config.captionstyles.keys())
            if btype == 'Equation':
                eq = self.selectEquation(blink)
                if eq is not None:
                    blink = eq
                else:
                    print("No equation node found")
                blabel = "equation:"+bname
            elif btype == 'Heading':
                blabel = self.config.style2heading[blink.getStyle()]+":"+bname
            elif btype == 'Footnote':
                blabel = "footnote:" + bname
            elif btype == 'Caption':
                blabel = "float:"+bname
            else:
                blabel = "none:"+bname
            self.config.bookmarks[blink] = {'type':btype, 'label':blabel, 'name':bname, 'id':bid}
            self.config.labels[bname] = blabel

            if self.verbose:
                print("Bookmark link: "+str(blink))
                print("Bookmark type: "+btype)

    def selectEquation(self, blink):
        prevEq, nextEq = self.findSurroundingEquations(blink)

        if prevEq is None and nextEq is not None:
            return nextEq
        if nextEq is None and prevEq is not None:
            return prevEq
        if nextEq is None and prevEq is None:
            return None

        self.xelatexcode.startToString()
        MathParagraphProcessor(self, prevEq)
        self.xelatexcode.endToString()
        prevEqCode = self.xelatexcode.getString()
        print()
        print("--------------")
        print("Equation 1:")
        print(prevEqCode)

        self.xelatexcode.startToString()
        MathParagraphProcessor(self, nextEq)
        self.xelatexcode.endToString()
        nextEqCode = self.xelatexcode.getString()
        print("--------------")
        print("Equation 2:")
        print(nextEqCode)

        print("--------------")
        res = input("Choose equation 1 or 2: ")
        if res == "1":
            return prevEq
        elif res =="2":
            return nextEq
        else:
            return None

    def findSurroundingEquations(self, blink):
        body = self.docx.documentxml.getBody()
        prevEq = None
        nextEq = None
        goToNext = False
        for mp in body.iter():
            if mp.tag == M+"oMathPara":
                if not goToNext:
                    prevEq = mp
                else:
                    nextEq = mp
                    break
            elif mp is blink:
                goToNext = True
        return (prevEq, nextEq)


class FieldCharProcessor(Processor):
    def process(self):
        if self.node.isEnd():
            self.config.suppressFieldContents = False
            return
        elif self.node.isSeparate():
            return

        fieldcodes = self.node.getInstructionText().split()

        if fieldcodes[0] == 'NOTEREF' or fieldcodes[0] == 'REF':
            refname = fieldcodes[1]
            if refname in self.config.labels:
                self.config.suppressFieldContents = True
                rtext = self.node.getResultText().split()
                for rt in rtext:
                    if rt in self.config.floatnames:
                        self.xelatexcode.append(rt+"~")
                        break
                self.xelatexcode.append("\\ref{"+self.config.labels[refname]+"}")
            else:
                print("Reference "+refname+" not found!")
        elif fieldcodes[0] == 'SEQ' and fieldcodes[1] in ['Equation', 'EquationNumber']:
            pass
        elif fieldcodes[0] == 'SEQ' and fieldcodes[1] in self.config.floatnames:
            pass
        # Specially added for Thesis Vis
        elif fieldcodes[0] == 'SEQ' and fieldcodes[1] in ['Example', 'example']:
            self.config.cancelcurrentcaption = True
        elif fieldcodes[0] == 'STYLEREF':
            pass
        elif fieldcodes[0] == 'ADDIN' and len(fieldcodes)>1:
            if fieldcodes[1] == 'REFMGR.CITE':
                pass
        else:
            print("!! Complex field code ignored: "+" ".join(fieldcodes))


class FieldSimpleProcessor(Processor):
    def process(self):
        fieldcodes = self.node.getInstructionText().split()

        if fieldcodes[0] == 'NOTEREF' or fieldcodes[0] == 'REF':
            refname = fieldcodes[1]
            if refname in self.config.labels:
                ### XXX Quick & Dirty !!!
                rtext = self.node.getContainedText().split()
                for rt in rtext:
                    if rt in self.config.floatnames:
                        self.xelatexcode.append(rt+"~")
                        break
                self.xelatexcode.append("\\ref{"+self.config.labels[refname]+"}")
            else:
                print("Reference "+refname+" not found!")
        elif fieldcodes[0] == 'SEQ' and fieldcodes[1] in ['Equation', 'EquationNumber']:
            pass
        elif fieldcodes[0] == 'SEQ' and fieldcodes[1] in self.config.floatnames:
            pass
        elif fieldcodes[0] == 'STYLEREF':
            pass
        elif fieldcodes[0] == 'ADDIN' and fieldcodes[1] == 'REFMGR.CITE':
            pass


class SubProcessor(Processor):
    def process(self):
        if len(self.node) > 0:
            node = self.node[0]
            while node is not None:
                if node.tag == W+'r':
                    RunProcessor(self, node)
                elif node.tag == W+'smartTagPr':
                    pass
                elif node.tag == W+'pPr':
                    pass
                elif node.tag == M+'oMathPara':
                    MathParagraphProcessor(self, node)
                elif node.tag == M+'oMath':
                    MathProcessor(self, node)
                elif node.tag == W+'bookmarkStart' or node.tag == W+'bookmarkEnd':
                    pass
                elif node.tag == W+'fldSimple':
                    FieldSimpleProcessor(self, node)
                elif node.tag == W+'smartTag':
                    SubProcessor(self, node)
                elif node.tag == W+'hyperlink':
                    SubProcessor(self, node)
                elif node.tag == W+'t':
                    TextProcessor(self, node)
                elif node.tag == W+'footnoteReference':
                    FootnoteProcessor(self, node)
                elif node.tag == W+'footnoteRef':
                    pass
                elif node.tag == M+'oMathPara':
                    MathParagraphProcessor(self, node)
                elif node.tag == W+'tab':
                    self.xelatexcode.append(" ")
                elif node.tag == W+'br':
                    self.xelatexcode.append(" ")
                elif node.tag == W+'softHyphen':
                    self.xelatexcode.append("\-") # Maybe not: this prevents hyphenation elsewhere in the word
                elif node.tag == W+'rPr' or node.tag == W+'lastRenderedPageBreak':
                    pass
                elif node.tag == W+'pict':
                    print("!! pict element not supported! Convert to pdf image.")
                elif node.tag == W+'drawing':
                    DrawingProcessor(self, node)
                elif node.tag == W+'fldChar':
                    FieldCharProcessor(self, node)
                elif node.tag == W+'instrText':
                    pass
                elif node.tag == W+'noBreakHyphen':
                    self.xelatexcode.append("{-}")
                else:
                    print("!! SmartTag subnode {:s} ignored".format(node.tag))
                node = node.getnext()


class CaptionProcessor(Processor):
    def __init__(self, parentprocessor, node, openbookmarks, lastdrawing=None):
        self.lastdrawing = lastdrawing
        self.openbookmarks = openbookmarks
        Processor.__init__(self, parentprocessor, node)

    def process(self):
        self.config.inCaption = True

        if self.node.isFirstWithStyle():
            self.config.buildCaption()
            if self.lastdrawing is None:
                self.config.currentcaptionp = self.node
            else:
                self.config.currentcaptionp = self.lastdrawing
            self.config.xelatexcode.startToString()

        self.processBookmarks()
        ParagraphProcessor(self, self.node)

        self.config.inCaption = False

    def processBookmarks(self):
        allbookmarks = {}
        for b in self.node.findAll('bookmarkStart'):
            allbookmarks[b.getBookmarkId()] = b
        allbookmarks.update(self.openbookmarks)

        for bid in allbookmarks.keys():
            b = allbookmarks[bid]
            btype, link = b.getBookmarkType(self.config.style2heading.keys(), self.config.floatnames, self.config.captionstyles.keys())
            bname = b.getBookmarkName()
            if btype == 'Caption' and bid not in self.config.currentcaptionlabels.keys():
                self.config.currentcaptionlabels[bid] = "float:"+bname


class TableCaptionProcessor(Processor):
    def __init__(self, parentprocessor, node, openbookmarks):
        self.openbookmarks = openbookmarks
        Processor.__init__(self, parentprocessor, node)

    def process(self):
        self.config.inCaption = True

        if self.node.isFirstWithStyle():
            self.config.buildTableCaption()
            self.config.currenttablecaptiontbl = self.findNextTable()
            self.config.xelatexcode.startToString()

        self.processBookmarks()
        ParagraphProcessor(self, self.node)

        self.config.inCaption = False

    def processBookmarks(self):
        allbookmarks = {}
        for b in self.node.findAll('bookmarkStart'):
            allbookmarks[b.getBookmarkId()] = b
        allbookmarks.update(self.openbookmarks)

        for bid in allbookmarks.keys():
            b = allbookmarks[bid]
            cstyles = list(self.config.captionstyles.keys())
            cstyles.extend(list(self.config.tablecaptionstyles.keys()))
            btype, link = b.getBookmarkType(self.config.style2heading.keys(), self.config.floatnames, cstyles)
            bname = b.getBookmarkName()
            if btype == 'Caption' and bid not in self.config.currentcaptionlabels.keys():
                self.config.currentcaptionlabels[bid] = "float:"+bname

    def findNextTable(self):
        nextn = self.node.getnext()
        if nextn.tag == W+'tbl':
            return nextn
        else:
            raise ValueError("Table note's next node is not a table!")


class TableNoteProcessor(Processor):
    def __init__(self, parentprocessor, node, lasttable):
        self.lasttable = lasttable
        print("Processing table note")
        print(self.lasttable)
        Processor.__init__(self, parentprocessor, node)

    def process(self):
        self.config.inTableNote = True

        if self.node.isFirstWithStyle():
            self.config.buildTableNote()
            self.config.currenttablenotetbl = self.lasttable
            self.config.xelatexcode.startToString()

        ParagraphProcessor(self, self.node)

        self.config.inTableNote = False


class AlternateContentProcessor(Processor):
    def __init__(self, parentprocessor, node):
        self.choices = []
        Processor.__init__(self, parentprocessor, node)

    def process(self):
        if len(self.node) > 0:
            node = self.node[0]
            while node is not None:
                if node.tag == MC+'Choice':
                    self.choices.append(node)
                elif node.tag == MC+'Fallback':
                    self.choices.append(node)
                else:
                    print("!! Alternate Content subnode {:s} ignored".format(node.tag))
                node = node.getnext()
        self.selectChoice()

    def processChoice(self, n):
        if len(n) > 0:
            node = n[0]
            while node is not None:
                if node.tag == W+"drawing":
                    DrawingProcessor(self, node)
                else:
                    print("!! Alternate Content subnode {:s} ignored".format(node.tag))
                node = node.getnext()

    def selectChoice(self):
        self.processChoice(self.choices[0])
        return

        print("")
        print("Alternate Contents:")
        print("-------------------")
        i = 1
        for c in self.choices:
            if c.tag == MC+'Choice':
                print("{:d}: Choice (requires {:s})".format(i, c.getDependency()))
            else:
                print("{:d}: Fallback".format(i))
            i+=1
        print("-------------------")
        sc = -1
        while sc <= 0 or sc > len(self.choices):
            sc = int(input("Select choice: "))

        self.processChoice(self.choices[sc-1])
