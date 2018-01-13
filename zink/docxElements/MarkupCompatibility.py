#    Copyright 2007-2018 Roel Zinkstok
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
from zink import docxBase


class MarkupCompatibilityElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'mc'

    def get(self, key, default=None):
        # The Markup Compatibility XML file does not use prefixes!
        return ET.ElementBase.get(self, key, default)


class AlternateContent(MarkupCompatibilityElement):
    """
    The AlternateContent element contains the full set of all possible markup alternatives. Each possible
    alternative is contained within either a Choice or Fallback child element of the AlternateContent element.
    """
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
