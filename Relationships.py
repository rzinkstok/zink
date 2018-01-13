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




