import os
import subprocess
from docxFrame import *
from UnicodeToLaTeXLibrary import *

class ConversionConfiguration: # ConversionTracker?
	def __init__(self):
		self.style2heading = {'Kop1':'chapter', 'Heading1':'chapter', 'Kop2':'section', 'Heading2':'section', 'Kop3':'subsection', 'Heading3':'subsection', 'Kop4':'subsubsection', 'Heading4':'subsubsection', 'Kop5':'paragraph', 'Heading5':'paragraph', 'Kop7':'subparagraph', 'Heading7':'subparagraph', 'Kop9':'sssssssSection', 'Heading9':'sssssssSection', 'ZinkHeading1':'chapter', 'ZinkHeading2':'section', 'ZinkHeading3':'subsection', 'ZinkHeading4':'subsubsection'}
		self.style2environment = {'quotation':'quotation', 'quote':'quotation', 'quotation_original':'quotation_original', 'quotation_translation':'quotation_translation', 'chapterquotation':'chapterquotation', 'chapterquote':'chapterquotation', 'blocktext':'zinkquotation', 'zinkquotation':'zinkquotation', 'zinkdescription':'zinkdescription', 'zinkequation':'zinkequation'}
		self.styleproperty2command = {'b':'textbf', 'i':'textit', 'u':'ul', 'superscript':'textsuperscript', 'subscript':'textsubscript', 'nor':'mathrm', 'double-struck':'mathbb'}
		self.floatnames = ['Figure', 'Figuur', 'Afbeelding', 'Tabel', 'Table', 'Note']
		self.captionstyles = {'Caption':'caption',
								'Bijschrift':'caption',
								'Figurenote':'caption'}
		self.tablecaptionstyles = {'TableCaption':'caption'}
		self.tablenotestyles = ['Tablenote', 'Tablenote0']
		self.mathfunctions = {	'sin':'\\sin', 
								'cos':'\\cos',
								'arccos':'\\arccos',
								'csc':'\\csc',
								'exp':'\\exp',
								'min':'\\min',
								'sinh':'\\sinh',
								'arcsin':'\\arcsin',
								'cosh':'\\cosh',
								'ln':'\\ln',
								'arctan':'\\arctan',
								'cot':'\\cot',
								'lim':'\\lim',
								'log':'\\log',
								'sec':'\\sec',
								'tan':'\\tan',
								'arg':'\\arg',
								'coth':'\\coth', 
								'dim':'\\dim', 
								'max':'\\max',
								'sin':'\\sin', 
								'tanh':'\\tanh'}
		self.mathaccents = {773:"\\overline"}
		self.convertable_image_extensions = ['.gif']
		self.incompatible_image_extensions = ['.emf', '.xml']
		
		self.useDirectStyling = True
		self.useTableColors = True
		
		self.supportedScriptLanguages = {'el-GR':'greek'}
		self.supportedFonts = {'Greek':'greek', 'Symbol':None}
		
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
		#self.currenttablecaptionlabels = {}
		#self.cancelcurrentcaption = False
		
		self.tablenotes = {}
		self.sortedtablenotekeys = []
		self.currenttablenotetbl = None
		self.currenttablenotelabels = {}
		
		self.colors = []
		
		self.xelatexcode = XeLaTeXCode()

	def newChapter(self):
		self.currentChapter += 1
		self.currentTable = 0
		self.xelatexcode.newChapter("Chapter{:d}".format(self.currentChapter))
	
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
			tcaption = "" #"\caption{"
			for l in self.currentcaptionlabels.values():
				tcaption += "\label{"+l+"}"
			tcaption += self.xelatexcode.getString()
			#tcaption += "}"
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


class XeLaTeXCode:
	def __init__(self):
		self.reset()

	def reset(self):
		self.data = {"Tables":"", "Frontmatter":""}
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
		
	def writeRootFile(self, basename, fnames):
		root = ""
		root += "\\documentclass[10pt]{memoir}\n"
		root += "\\usepackage{zinkLood}\n"
		root += "\\begin{document}\n"
		
		root += self.getColorDefinitions()
		
		for f in fnames:
			# get filename (no path) without extension
			fn = os.path.splitext(os.path.split(f)[-1])[0]
			root += "\\include{" + fn + "}\n"
		
		root += "\\end{document}\n"
		
		rfn = basename+"_root.tex"
		print("Writing root file {:s}".format(rfn))
		fp = open(rfn, "w")
		fp.write(root)
		fp.close()
		
	def toFile(self, basename):
		fnames = []
		for c in self.chapters:
			if c == "Tables" and self.data[c] == "":
				continue
			fname = basename + "_" + c + ".tex"
			fnames.append(fname)
			print("Writing {:s}".format(fname))
			fp = open(fname, "w")
			fp.write(self.data[c])
			fp.close()
		self.writeRootFile(basename, fnames)
			
		
class docx2xelatex:
	def __init__(self, docx):
		self.docx = docx
		self.docxfilename = self.docx.filename
		self.basefilename = os.path.splitext(self.docxfilename)[0]
		
		self.config = ConversionConfiguration()
		self.xelatexcode = self.config.xelatexcode
		
		print()
		print("Converting {:s} to xelatex".format(self.docxfilename))
		print("Base filename: {:s}".format(self.basefilename))
		print("Path: {:s}".format(self.docx.path))
		print()
		
		self.docx.documentxml.save(filename=os.path.join(self.docx.path, 'document_orig.xml'), pretty_print=True)
		if self.docx.footnotesxml is not None:
			self.docx.footnotesxml.save(filename=os.path.join(self.docx.path, 'footnotes_orig.xml'), pretty_print=True)
	
	def buildReferenceList(self):
		itexts = []
		print()
		print("==============================")
		print("  Building list of references")
		print("==============================")
		print()
		print("Document body:")
		d_itexts = self.findFieldInstructions(self.docx.documentxml.getBody())
		itexts.extend(i for i in d_itexts if i not in itexts)
		if self.docx.footnotesxml is not None:
			print("Document footnotes:")
			f_itexts = self.findFieldInstructions(self.docx.footnotesxml.getFootnotes())
			itexts.extend(i for i in f_itexts if i not in itexts)
		
		self.config.references = []
		for i in itexts:
			fieldcodes = i.split()
			if fieldcodes[0] == "REF" or fieldcodes[0] == "NOTEREF" or fieldcodes[0] == "PAGEREF" and fieldcodes[1].find("_Toc")<0:
				self.config.references.append(fieldcodes[1])
		#print(str(self.config.references))
		print()
		print("Found {:d} references".format(len(self.config.references)))
		print()
		input()
	
	def buildBookmarks(self):
		print()
		print("===============================")
		print("  Building list of bookmarks")
		print("===============================")
		print()
		bms = self.docx.documentxml.getBody().findAll('bookmarkStart')
		print()
		print("Found {:d} bookmarks".format(len(bms)))
		print()
		for bm in bms:
			BookmarkProcessor(self, bm)
		print()
		print("Found {:d} referenced bookmarks".format(len(self.config.bookmarks)))	
		#print(str(self.config.bookmarks))
		print()
		input()
	
	def buildCaptions(self):
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
			
		print()
		print("Found {:d} captions".format(len(self.config.captions)))
		print()
		input()
	
	def buildTableCaptions(self):
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
			
		print()
		print("Found {:d} table captions".format(len(self.config.tablecaptions)))
		print()
		input()
		
	def buildTableNotes(self):
		print()
		print("==============================")
		print("  Building list of table notes")
		print("==============================")
		print()
		#lasttable = None
		
		for n in self.docx.getDocumentBody().iter():
			if n.tag == W+'p' and n.isTableNote(self.config.tablenotestyles):
				if n.getprevious().tag == W+'tbl':
					TableNoteProcessor(self, n, n.getprevious())
			elif n.tag == W+'tbl':
				pass
			else:
				#if lasttable is not None:
				#	print("End table")
				#	lasttable = None
				pass
		
		if self.config.currenttablenotetbl is not None:
			self.config.buildTableNote()
			self.config.currenttablenotetbl = None
			self.config.currenttablenotelabels = {}
		
		print()
		print("Found {:d} table notes".format(len(self.config.tablenotes)))
		print()
		input()
		
	def findFieldInstructions(self, node):
		itexts = []
		flds = node.findAll('fldChar')
		fflds = [f for f in flds if f.isBegin()] 
		print("Found {:d} complex fields".format(len(fflds)))
		for f in fflds:
			itexts.append(f.getInstructionText())
			
		sflds = node.findAll('fldSimple')
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
		self.xelatexcode.toFile(self.basefilename)
		
class ParagraphProcessor:
	def __init__(self, parentprocessor, p):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.p = p
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def startHeading(self):
		h = self.config.style2heading[self.p.getStyle()]
		self.config.inHeading = True
		if h == 'chapter':
			self.config.newChapter()
			print()
			print("New chapter: {:s}".format(self.p.getText()))
			print()
		self.xelatexcode.addNL()
		self.xelatexcode.append('\\'+h+'{')
		
	def endHeading(self):
		self.xelatexcode.append('}')
		self.xelatexcode.addNL()
		self.config.inHeading = False
		if self.p in self.config.bookmarks:
			h = self.config.style2heading[self.p.getStyle()]
			self.xelatexcode.append("\\label{"+h+":"+self.config.bookmarks[self.p]["name"]+"}")
			self.xelatexcode.addNL()
		
	def startList(self):
		numId, iLvl = self.p.getNumbering()
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
		numId, iLvl = self.p.getNumbering()
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
		if self.p.isFirstListItem():
			self.startList()
		self.xelatexcode.addNL()
		self.xelatexcode.append('\\item ')
		
	def endListItem(self):
		if self.p.isLastListItem():
			self.endList()
		
	def startEnvironmentIfNeeded(self):
		if self.p.isFirstWithEnvironment():
			self.xelatexcode.append("\\begin{" + self.p.getEnvironment(self.config.style2environment) + "}")
	
	def endEnvironmentIfNeeded(self):
		if self.p.isLastWithEnvironment():
			self.xelatexcode.append("\\end{" + self.p.getEnvironment(self.config.style2environment) + "}")
		
	def process(self):
		# Check heading, list, caption, footnote
		if self.p.isHeading(self.config.style2heading):
			if self.p.hasText():
				self.startHeading()
		elif self.p.isListItem():
			self.startListItem()
		elif not self.config.inCaption and self.p.isCaption(self.config.captionstyles):
			if self.p in self.config.captions.keys():
				self.config.xelatexcode.append(self.config.captions[self.p])
				self.xelatexcode.addNL()
				del(self.config.captions[self.p])
			return
		elif not self.config.inCaption and self.p.isCaption(self.config.tablecaptionstyles):
			return
				
		if self.p.isEnvironment(self.config.style2environment):
			self.startEnvironmentIfNeeded()
		
		if len(self.p) > 0:
			node = self.p[0]
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
				
		if self.p.isEnvironment(self.config.style2environment):
			self.endEnvironmentIfNeeded()
		
		if self.p.isHeading(self.config.style2heading):
			if self.p.hasText():
				self.endHeading()
		elif self.p.isListItem():
			self.endListItem()
		
		if not self.config.inFootnote and not self.config.inTable and not self.config.inCaption and not self.config.inTableNote:
			self.xelatexcode.addNL()
			self.xelatexcode.addNL()
		
class RunProcessor:
	def __init__(self, parentprocessor, r):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.r = r
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
	
	def startStyleIfNeeded(self, s):
		if self.r.isFirstWithStyleProperty(s):
			self.xelatexcode.append('\\' + self.config.styleproperty2command[s] + '{')
		
	def endStyleIfNeeded(self, s):
		if self.r.isLastWithStyleProperty(s):
			self.xelatexcode.append('}')
		
	def startStyles(self):
		if self.r.isBold(self.config.useDirectStyling):
			self.startStyleIfNeeded('b')
		if self.r.isItalic(self.config.useDirectStyling):
			self.startStyleIfNeeded('i')
		if self.r.isUnderline(self.config.useDirectStyling):
			self.startStyleIfNeeded('u')
		if self.r.isSuperScript():
			self.startStyleIfNeeded('superscript')
		if self.r.isSubScript():
			self.startStyleIfNeeded('subscript')
		#print("\n!! Run properties subnode {:s} ignored\n".format(node.tag))
			
	def endStyles(self):
		if self.r.isBold(self.config.useDirectStyling):
			self.endStyleIfNeeded('b')
		if self.r.isItalic(self.config.useDirectStyling):
			self.endStyleIfNeeded('i')
		if self.r.isUnderline(self.config.useDirectStyling):
			self.endStyleIfNeeded('u')
		if self.r.isSuperScript():
			self.endStyleIfNeeded('superscript')
		if self.r.isSubScript():
			self.endStyleIfNeeded('subscript')
	
	def processLanguage(self):
		language = self.r.getLanguage()
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
		font = self.r.getFont()
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
		
		if len(self.r) > 0:
			node = self.r[0]
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

class TextProcessor:
	def __init__(self, parentprocessor, t):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.t = t
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		LaTeXizer(self, self.t.getText())
		
class LaTeXizer:
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

class SymProcessor:
	def __init__(self, parentprocessor, s):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.s = s
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		c = chr(self.s.getSymbolCode())
		f = self.s.getSymbolFont()
		uc = self.config.fontconvertor.getUnicode(f, c)
		l = LaTeXizer(self, uc, isString=True)
		s = l.getXeLaTeXString()
		print("Sym: {:s}".format(s))
		self.xelatexcode.append(s)
		
		
class FootnoteProcessor:
	def __init__(self, parentprocessor, f):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.f = f
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):		
		fid = self.f.getFootnoteId()
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
		
class TableProcessor:
	def __init__(self, parentprocessor, t):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.t = t
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
	
	def arabicToAlphabetic(self, a):
		b = 0
		while a>26:
			a -= 26
			b += 1
		
		if b>0:
			firstletter = chr(ord("@")+b)
		else:
			firstletter = ""
		return firstletter + chr(ord("@")+a)
				
	def process(self):
		suppressTable = not self.t.hasText('()0123456789', ['Caption', 'Figurenote', 'Tablenote', 'Tablenote0'])
		
		self.config.startTable(suppressTable)
		
		if not suppressTable:
			tableid = "tabc"+self.arabicToAlphabetic(self.config.currentChapter)+"t"+self.arabicToAlphabetic(self.config.currentTable) # ord("@") = ord("A") - 1
			if self.t in self.config.tablenotes:
				tablenotes = self.config.tablenotes[self.t]
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
			
			tw = self.t.getWidth()
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
				self.xelatexcode.append("\tdoinside=\\tablefont\\tiny,%")
				self.xelatexcode.addNL()
				self.xelatexcode.append("\tcaption={")
				if self.t in self.config.tablecaptions:
					self.xelatexcode.append(self.config.tablecaptions[self.t])
					del(self.config.tablecaptions[self.t])
				self.xelatexcode.append("}%")
				self.xelatexcode.addNL()
				self.xelatexcode.append("]%")
				self.xelatexcode.addNL()
				self.xelatexcode.append("{%")
				self.xelatexcode.addNL()
			
			ncols = self.t.getNumberOfLogicalColumns()
			nrows = self.t.getNumberOfRows()
			
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
		rows = self.t.getRows()
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
		
		
class MathParagraphProcessor:
	def __init__(self, parentprocessor, mp):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.mp = mp
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		self.config.inMathEnvironment = True
		
		self.xelatexcode.addNL()
		self.xelatexcode.addNL()
		self.xelatexcode.append("\\begin{equation}")
		self.xelatexcode.addNL()
		
		if self.mp in self.config.bookmarks:
			self.xelatexcode.append("\\label{equation:"+self.config.bookmarks[self.mp]["name"]+"}")
			self.xelatexcode.addNL()
		
		node = self.mp[0]
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

class MathProcessor:
	def __init__(self, parentprocessor, m):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.m = m
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		if not self.config.inMathEnvironment:
			self.config.inMath = True
		
		mlc = self.processMathElement(self.m)
		
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

class MathRunProcessor:
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
		if self.r.isRoman():
			self.startStyleIfNeeded('nor')
		if self.r.isDoubleStruck():
			self.startStyleIfNeeded('double-struck')
			
	def startStyleIfNeeded(self, s):
		if self.r.isFirstWithStyleProperty(s):
			self.xelatexstring += "\\" + self.config.styleproperty2command[s] + "{"

	def endStyles(self):
		if self.r.isRoman():
			self.endStyleIfNeeded('nor')
		if self.r.isDoubleStruck():
			self.endStyleIfNeeded('double-struck')
		
	def endStyleIfNeeded(self, s):
		if self.r.isLastWithStyleProperty(s):
			self.xelatexstring += "}"
			
class MathTextProcessor:
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
		
class DrawingProcessor:
	def __init__(self, parentprocessor, d):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.d = d
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		commented = False
		latexname = None
		
		refID = self.d.getImageReferenceId()
		if refID is None:
			print("Skipping drawing (no reference ID found)")
			self.xelatexcode.addNL()
			#self.xelatexcode.append("%% Drawing omitted")
			self.xelatexcode.addNL()
			#return	
		else:
			sx, sy = self.d.getImageSize()
		
			relxml = self.docx.getRelationsXML()
			zippath = os.path.join('word', relxml.resolveRelation(refID))
		
			filename = os.path.split(zippath)[-1]
		
			mediadirname = "media"
			mediapath = os.path.join(self.docx.path, "media")
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
		if self.d in self.config.captions:
			self.xelatexcode.append("\\caption{")
			self.xelatexcode.append(self.config.captions[self.d])
			self.xelatexcode.append("}")
			self.xelatexcode.addNL()
			del(self.config.captions[self.d])
			
		if commented:
			self.xelatexcode.append("%")
		self.xelatexcode.append("\\end{")
		self.xelatexcode.append("figure")
		self.xelatexcode.append("}")
		self.xelatexcode.addNL()

class BookmarkProcessor:
	def __init__(self, parentprocessor, b):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.b = b
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		bid = self.b.getBookmarkId()
		bname = self.b.getBookmarkName()
		if bname in self.config.references:
			print("=========================")
			print("Bookmark id: "+bid+"; name: "+bname)
			btype, blink = self.b.getBookmarkType(self.config.style2heading.keys(), self.config.floatnames, self.config.captionstyles.keys())
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
		
class FieldCharProcessor:
	def __init__(self, parentprocessor, fc):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.fc = fc
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		if self.fc.isEnd():
			self.config.suppressFieldContents = False
			return
		elif self.fc.isSeparate():
			return

		fieldcodes = self.fc.getInstructionText().split()
		
		if fieldcodes[0] == 'NOTEREF' or fieldcodes[0] == 'REF':
			refname = fieldcodes[1]
			if refname in self.config.labels:
				self.config.suppressFieldContents = True
				rtext = self.fc.getResultText().split()
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
			
			
class FieldSimpleProcessor:
	def __init__(self, parentprocessor, fs):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.fs = fs
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		fieldcodes = self.fs.getInstructionText().split()
		
		if fieldcodes[0] == 'NOTEREF' or fieldcodes[0] == 'REF':
			refname = fieldcodes[1]
			if refname in self.config.labels:
				### XXX Quick & Dirty !!!
				rtext = self.fs.getContainedText().split()
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
	
class SubProcessor:
	def __init__(self, parentprocessor, n):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.n = n
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		if len(self.n) > 0:
			node = self.n[0]
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
		
class CaptionProcessor:
	def __init__(self, parentprocessor, cp, openbookmarks, lastdrawing=None):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.cp = cp
		self.lastdrawing = lastdrawing
		self.openbookmarks = openbookmarks
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		self.config.inCaption = True
		
		if self.cp.isFirstWithStyle():
			self.config.buildCaption()
			if self.lastdrawing is None:
				self.config.currentcaptionp = self.cp
			else:
				self.config.currentcaptionp = self.lastdrawing
			self.config.xelatexcode.startToString()

		self.processBookmarks()	
		ParagraphProcessor(self, self.cp)
		
		self.config.inCaption = False
			
	def processBookmarks(self):
		allbookmarks = {}
		for b in self.cp.findAll('bookmarkStart'):
			allbookmarks[b.getBookmarkId()] = b
		allbookmarks.update(self.openbookmarks)
		
		for bid in allbookmarks.keys():
			b = allbookmarks[bid]
			btype, link = b.getBookmarkType(self.config.style2heading.keys(), self.config.floatnames, self.config.captionstyles.keys())
			bname = b.getBookmarkName()
			if btype == 'Caption' and bid not in self.config.currentcaptionlabels.keys():
				self.config.currentcaptionlabels[bid] = "float:"+bname

class TableCaptionProcessor:
	def __init__(self, parentprocessor, tcp, openbookmarks):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.tcp = tcp
		self.openbookmarks = openbookmarks
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
		
	def process(self):
		self.config.inCaption = True
		
		if self.tcp.isFirstWithStyle():
			self.config.buildTableCaption()
			self.config.currenttablecaptiontbl = self.findNextTable()
			self.config.xelatexcode.startToString()

		self.processBookmarks()	
		ParagraphProcessor(self, self.tcp)
		
		self.config.inCaption = False
		
	def processBookmarks(self):
		allbookmarks = {}
		for b in self.tcp.findAll('bookmarkStart'):
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
		nextn = self.tcp.getnext()
		if nextn.tag == W+'tbl':
			return nextn
		else:
			raise ValueError("Table note's next node is not a table!")
			
class TableNoteProcessor:
	def __init__(self, parentprocessor, tn, lasttable):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.tn = tn
		self.lasttable = lasttable
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		print("Processing table note")
		print(self.lasttable)
		self.process()
		
	def process(self):
		self.config.inTableNote = True
		
		if self.tn.isFirstWithStyle():
			self.config.buildTableNote()
			self.config.currenttablenotetbl = self.lasttable
			self.config.xelatexcode.startToString()
			
		ParagraphProcessor(self, self.tn)
		
		self.config.inTableNote = False
		
	
class AlternateContentProcessor:
	def __init__(self, parentprocessor, ac):
		self.parentprocessor = parentprocessor
		self.config = self.parentprocessor.config
		self.ac = ac
		self.choices = []
		self.docx = self.parentprocessor.docx
		self.xelatexcode = self.parentprocessor.xelatexcode
		self.process()
	
	def process(self):
		if len(self.ac) > 0:
			node = self.ac[0]
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