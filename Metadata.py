import sys
import os
import shutil
import subprocess
import zipfile
import codecs
import unittest
from lxml import etree as ET
import docxBase


class CorePropertiesElement(docxBase.OfficeOpenXMLElement):
	def _init(self):
		self.nsprefix = 'cp'
			
class DCElement(docxBase.OfficeOpenXMLElement):
	def _init(self):
		self.nsprefix = 'dc'
				
class coreProperties(CorePropertiesElement):
	pass
	
class creator(DCElement):
	pass
	


CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
CP = "{%s}" % CP_NS
docxBase.NSMAP['cp'] = CP_NS

DC_NS = "http://purl.org/dc/elements/1.1/" 
DC = "{%s}" % DC_NS
docxBase.NSMAP['dc'] = DC_NS

#xmlns:dcterms="http://purl.org/dc/terms/" 
#xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
#xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"



namespace = docxBase.lookup.get_namespace(CP_NS)
namespace.update(vars())


