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


from zink import docxBase


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

namespace = docxBase.lookup.get_namespace(CP_NS)
namespace.update(vars())


