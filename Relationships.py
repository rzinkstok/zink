###!/Library/Frameworks/Python.framework/Versions/3.1/bin/python

from lxml import etree as ET
import docxBase


class PackageRelationshipsElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'rel'

    def get(self, key, default=None):
        # The Relationships XML file does not use prefixes!
        return ET.ElementBase.get(self, key, default)


class Relationships(PackageRelationshipsElement):
    pass


class Relationship(PackageRelationshipsElement):
    pass


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = "{%s}" % REL_NS
docxBase.NSMAP['rel'] = REL_NS

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = "{%s}" % R_NS
docxBase.NSMAP['r'] = R_NS

namespace = docxBase.lookup.get_namespace(REL_NS)
namespace.update(vars())




