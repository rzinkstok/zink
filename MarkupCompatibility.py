#!/Library/Frameworks/Python.framework/Versions/3.1/bin/python
import sys
import os


from lxml import etree as ET
import docxBase

class MarkupCompatibilityElement(docxBase.OfficeOpenXMLElement):
	def _init(self):
		self.nsprefix = 'mc'
	
	def get(self, key, default=None):
		# The Markup Compatibility XML file does not use prefixes!
		return ET.ElementBase.get(self, key, default)

class AlternateContent(MarkupCompatibilityElement):
	# The AlternateContent element contains the full set of all possible markup alternatives. Each possible
	# alternative is contained within either a Choice or Fallback child element of the AlternateContent element.
	pass
	
class Choice(MarkupCompatibilityElement):
	def getDependency(self):
		return self.get('Requires')

class Fallback(MarkupCompatibilityElement):
	pass

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MC = "{%s}" % MC_NS
docxBase.NSMAP['mc'] = MC_NS

namespace = docxBase.lookup.get_namespace(MC_NS)
namespace.update(vars())
