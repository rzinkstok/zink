###!/Library/Frameworks/Python.framework/Versions/3.1/bin/python

import os
from lxml import etree as ET
import unicodedata as ucd

class UnicodeToLaTeXLibraryError(Exception):
	def __init__(self, codepoint):
		self.codepoint = codepoint
	
	def __str__(self):
		s = "Unknown codepoint {:d}".format(self.codepoint)
		return s

class FontEncodingConvertorError(Exception):
	def __init__(self, codepoint, font):
		self.codepoint = codepoint
		self.font = font
	
	def __str__(self):
		s = "Unknown codepoint {:d} for font {:s}".format(self.codepoint, self.font)
		return s

class UnicodeToLaTeXLibrary:
	def __init__(self, libfilename=None):
		self.latexchardict = {}
		self.xelatexchardict = {}
		self.latexmathchardict = {}
		if libfilename is None:
			libfilename = '/Users/rzinkstok/Documents/Development/zink/UnicodeToLaTeXLibrary.xml'
		self.libraryfilename = libfilename
		self.loadXMLToDict()
	
	def loadXMLToDict(self):
		self.latexchardict = {}
		self.xelatexchardict = {}
		self.latexmathchardict = {}
		xml = self.openXML(self.libraryfilename)
		
		# latex chars
		chars = xml.findall('latex-char')
		
		for char in chars:
			cp = int(char.findtext('codepoint'))
			lc = char.findtext('latexcode')
			self.latexchardict[cp] = lc
			name = ucd.name(chr(cp))
			print("Loading latex character {:s} ({:d} {:s})".format(lc, cp, name))
		
		# xelatex chars
		chars = xml.findall('xelatex-char')
		
		for char in chars:
			cp = int(char.findtext('codepoint'))
			lc = char.findtext('xelatexcode')
			self.xelatexchardict[cp] = lc
			name = ucd.name(chr(cp))
			print("Loading xelatex character {:s} ({:d} {:s})".format(lc, cp, name))
		
		# latex math chars
		chars = xml.findall('latex-math-char')
		
		for char in chars:
			cp = int(char.findtext('codepoint'))
			lc = char.findtext('latexmathcode')
			self.latexmathchardict[cp] = lc
			name = ucd.name(chr(cp))
			print("Loading latex math character {:s} ({:d} {:s})".format(lc, cp, name))
		
	def openXML(self, filename, mode='r'):
		if os.path.splitext(filename)[1] != '.xml':
			raise ValueError('File is not an xml file!')
		fp = open(filename, mode)
		xml = ET.parse(fp)
		fp.close()
		return xml
	
	def buildXML(self):
		print() 
		print("Building XML file from database...")
		print("----------------------------------")
		print()
		
		characters = ET.Element('characters')
				
		for cp in sorted(self.latexchardict.keys()):
			lc = self.latexchardict[cp]
			name = ucd.name(chr(cp))
			print("Adding latex character {:s} ({:d} {:s})".format(lc, cp, name))
			char = ET.SubElement(characters, 'latex-char')
			charcp = ET.SubElement(char, 'codepoint')
			charcp.text = "{:d}".format(cp)
			latexcode = ET.SubElement(char, 'latexcode')
			latexcode.text = "{:s}".format(lc)
			n = ET.SubElement(char, 'name')
			n.text = name
			
		for cp in sorted(self.xelatexchardict.keys()):
			lc = self.xelatexchardict[cp]
			name = ucd.name(chr(cp))
			print("Adding xelatex character {:s} ({:d} {:s})".format(lc, cp, name))
			char = ET.SubElement(characters, 'xelatex-char')
			charcp = ET.SubElement(char, 'codepoint')
			charcp.text = "{:d}".format(cp)
			latexcode = ET.SubElement(char, 'xelatexcode')
			latexcode.text = "{:s}".format(lc)
			n = ET.SubElement(char, 'name')
			n.text = name
			
		for cp in sorted(self.latexmathchardict.keys()):
			lc = self.latexmathchardict[cp]
			name = ucd.name(chr(cp))
			print("Adding latex math character {:s} ({:d} {:s})".format(lc, cp, name))
			char = ET.SubElement(characters, 'latex-math-char')
			charcp = ET.SubElement(char, 'codepoint')
			charcp.text = "{:d}".format(cp)
			latexcode = ET.SubElement(char, 'latexmathcode')
			latexcode.text = "{:s}".format(lc)
			n = ET.SubElement(char, 'name')
			n.text = name
			
		fp = open(self.libraryfilename, 'w')
		fp.write(ET.tostring(characters, encoding=str, pretty_print=True))
		fp.close()
	
	def insertLaTeXChar(self, c):
		codepoint = ord(c)
		lc = input("Enter LaTeX code for unicode character {:d}: ".format(codepoint))
		
		if lc == "xxx":
			raise IOError
		self.latexchardict[codepoint] = lc
		self.buildXML()
	
	def insertXeLaTeXChar(self, c):
		codepoint = ord(c)
		xc = input("Enter XeLaTeX code for unicode character {:d}: ".format(codepoint))
		
		if xc == "xxx":
			raise IOError
		self.xelatexchardict[codepoint] = xc
		self.buildXML()
		
	def insertLaTeXMathChar(self, c):
		codepoint = ord(c)
		lc = input("Enter LaTeX math code for unicode character {:d}: ".format(codepoint))
		
		if lc == "xxx":
			raise IOError
		self.latexmathchardict[codepoint] = lc
		self.buildXML()
		
	def getLaTeXChar(self, c):
		codepoint = ord(c)
		if codepoint in self.latexchardict:
			return self.latexchardict[codepoint]
		else:
			raise UnicodeToLaTeXLibraryError(codepoint)
			
	def getXeLaTeXChar(self, c):
		codepoint = ord(c)
		if codepoint in self.xelatexchardict:
			return self.xelatexchardict[codepoint]
		else:
			raise UnicodeToLaTeXLibraryError(codepoint)
			
	def getLaTeXMathChar(self, c):
		codepoint = ord(c)
		if codepoint in self.latexmathchardict:
			return self.latexmathchardict[codepoint]
		else:
			raise UnicodeToLaTeXLibraryError(codepoint)
			


class FontEncodingConvertor:
	def __init__(self, mapfilename=None):
		if mapfilename is None:
			mapfilename = '/Users/rzinkstok/Documents/Development/zink/FontEncodingMap.xml'
		self.mapfilename = mapfilename
		self.buildMap()
	
	def buildMap(self):
		self.map = {}
		xml = self.openXML(self.mapfilename)
		fontnodes = xml.findall('font')
		
		for font in fontnodes:
			fontname = font.attrib['name']
			print(fontname)
			self.map[fontname] = {}
			for cp in font:
				origcp = int(cp.findtext('original'))
				print(origcp)
				unicodecp = int(cp.findtext('unicode'))
				self.map[fontname][origcp] = unicodecp
		print(self.map)
		
	def openXML(self, filename, mode='r'):
		if os.path.splitext(filename)[1] != '.xml':
			raise ValueError('File is not an xml file!')
		fp = open(filename, mode)
		xml = ET.parse(fp)
		fp.close()
		return xml
		
	def getUnicode(self, font, c):
		if font not in self.map:
			raise KeyError("font <%s> not found!" % font)
		codepoint = ord(c)
		if codepoint in self.map[font]:
			return chr(self.map[font][codepoint])
		else:
			raise FontEncodingConvertorError(codepoint, font)