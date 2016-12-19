from docxFrame import *
from UnicodeToLaTeXLibrary import *

fec = FontEncodingConvertor()

def processP(p):
	greektext = ''
	for node in p:
		if node.tag == W+'r':
			if node.hasProperty('rFonts', {W+"ascii":"Greek", W+"hAnsi":"Greek"}):
				greektext += node.getText()
	if len(greektext) > 0:
		print(greektext)
		print()
		for c in greektext:
			print(c, " -> ", ord(c), "%x" % ord(c), fec.getUnicode('greek', c))
		print()
		print()
		
fn = "/Users/rzinkstok/Dropbox/Zink/AA Current Work/20120825_Betti/Conversie/Logik.docx"
docx = DocX(fn)

node = docx.documentxml.getBody()[0]



while node is not None:
	if node.tag == W+'p':
		processP(node)
	node = node.getnext()
	