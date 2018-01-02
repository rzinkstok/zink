###!/Library/Frameworks/Python.framework/Versions/3.1/bin/python

import sys
import os
import shutil
import subprocess
import zipfile
import codecs
import unittest
from lxml import etree as ET
import docxBase


class DrawingMLElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'a'

class DrawingElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'a14'

class blip(DrawingMLElement):
    """
    This element specifies the existence of an image (binary large image or picture) and contains a reference to the
    image data.

    Parent elements: blipFill, buBlip
    Child elements: alphaBiLevel, alphaCeiling, alphaFloor, alphaInv, alphaMod, alphaModFix, alphaRepl,
                    biLevel, blur, clrChange, clrRepl, duotone, extLst, fillOverlay, grayscl, hsl, lum, tint
    """
    pass


class DrawingMLWordprocessingDrawingElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'wp'


class extent(DrawingMLWordprocessingDrawingElement):
    """
    This element specifies the extents of the parent DrawingML object within the document
    (i.e. its final height and width).

    Parent elements: anchor, inline

    Child elements: cx, cy
    """
    pass


class DrawingMLPictureElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'pic'


class chart(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'c'

    def getRelationId(self):
        return self.get('r:id')


A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
A = "{%s}" % A_NS
docxBase.NSMAP['a'] = A_NS

A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
A14 = "{%s}" % A14_NS
docxBase.NSMAP['a14'] = A14_NS

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WP = "{%s}" % WP_NS
docxBase.NSMAP['wp'] = WP_NS

PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PIC = "{%s}" % PIC_NS
docxBase.NSMAP['pic'] = PIC_NS

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
C = "{%s}" % C_NS
docxBase.NSMAP['c'] = C_NS

namespace = docxBase.lookup.get_namespace(A_NS)
namespace['blip'] = blip
namespace[None] = DrawingMLElement

namespace = docxBase.lookup.get_namespace(A14_NS)
namespace[None] = DrawingElement

namespace = docxBase.lookup.get_namespace(WP_NS)
namespace['extent'] = extent
namespace[None] = DrawingMLWordprocessingDrawingElement

namespace = docxBase.lookup.get_namespace(PIC_NS)
namespace[None] = DrawingMLPictureElement

namespace = docxBase.lookup.get_namespace(C_NS)
namespace[None] = DrawingMLElement
