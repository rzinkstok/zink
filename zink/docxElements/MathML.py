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


class MathMLElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'm'


class oMathPara(MathMLElement):
    """
    This element specifies a math paragraph, one or more display equations within a single paragraph.

    Parent elements: body, comment, customXml, deg, del, den, docPartBody, e, endnote, fldSimple,
    fName, footnote, ftr, hdr, hyperlink, ins, lim, moveFrom, moveTo, num, oMath, p, rt, rubyBase,
    sdtContent, smartTag, sub, sup, tbl, tc, tr, txbxContent

    Child elements: oMath, oMathParaPr
    """

    pass


class oMath(MathMLElement):
    """
    This element specifies an equation or mathematical expression. All equations are surrounded by oMath tags.

    Parent elements: body, comment, customXml, deg, del, den, docPartBody, e, endnote, fldSimple,
    fName, footnote, ftr, hdr, hyperlink, ins, lim, moveFrom, moveTo, num, oMath, p, rt, rubyBase,
    sdtContent, smartTag, sub, sup, tbl, tc, tr, txbxContent

    Child elements: acc, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd,
    customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart, d, del, eqArr,
    f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class sSub(MathMLElement):
    """
    This element specifies the subscript function sSub, which consists of a base e and a reduced-size
    scr placed below and to the right.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: e, sSubPr, sub
    """

    pass


class sSubSup(MathMLElement):
    """
    This element specifies the sub-superscript function, which consists of a base e, a reduced-size
    scr placed below and to the right, and a reduced-size scr placed above and to the right

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: e, sSubSupPr, sub, sup
    """

    pass


class sSup(MathMLElement):
    """
    This element specifies the subscript function sSub, which consists of a base e and a reduced-size
    scr placed below and to the right

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: e, sSupPr, sup
    """

    pass


class e(MathMLElement):
    """
    This element specifies the base argument of a mathematical function.

    Parent elements: acc, bar, borderBox, box, d, eqArr, func, groupChr, limLow, limUpp, mr, nary,
    phant, rad, sPre, sSub, sSubSup, sSup

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class sub(MathMLElement):
    """
    This element specifies the subscript of the Pre-Sub-Superscript function sPre.
    The element also specifies the subscript of the sSub en sSubSup functions

    Parent elements: nary, sPre, sSub, sSubSup

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class sup(MathMLElement):
    """
    This element specifies the superscript of the superscript function sSup.

    Parent elements: nary, sPre, sSubSup, sSup

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class d(MathMLElement):
    """
    This element specifies the delimiter function, consisting of opening and closing delimiters (such as
    parentheses, braces, brackets, and vertical bars), and an element contained inside. The delimiter
    may have more than one element, with a designated separator character between each element.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: dPr, e
    """

    pass


class dPr(MathMLElement):
    """
    This element specifies the properties of d, including the enclosing and separating characters, and the properties
    that affect the shape of the delimiters.

    Parent element: d

    Child elements: begChr, ctrlPr, endChr, grow, sepChr, shp
    """

    pass

class begChr(MathMLElement):
    """
    This element specifies the beginning, or opening, delimiter character. Mathematical delimiters
    are enclosing characters such as parentheses, brackets, and braces. If this element is omitted,
    the default begChr is '('.

    Parent element: dPr
    """

    pass


class endChr(MathMLElement):
    """
    This element specifies the ending, or closing, delimiter character. Mathematical delimiters
    are enclosing characters such as parentheses, brackets, and braces. If this element is omitted,
    the default endChr is ')'.

    Parent element: dPr
    """

    pass


class sepChr(MathMLElement):
    """
    This element specifies the character that separates base arguments e in the delimiter object
    d. If this element is omitted, the default sepChr is '|'.

    Parent element: dPr
    """

    pass


class nary(MathMLElement):
    """
    This element specifies an n-ary object, consisting of an n-ary object, a base (or operand),
    and optional upper and lower limits.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: e, naryPr, sub, sup
    """

    pass


class naryPr(MathMLElement):
    """
    This element specifies the properties of the n-ary object, including the type of n-ary operator
    that is used, the shape and height of the operator, the location of limits, and whether limits
    are shown or hidden.

    Parent element: nary

    Child elements: chr, ctrlPr, grow, limLoc, subHide, supHide
    """

    pass


class chr(MathMLElement):
    """
    This element specifies the type of combining diacritical mark attached to the base of the
    accent function. If this property is omitted, the default accent character is U+0302.

    Parent elements: accPr, groupChrPr, naryPr
    """

    pass


class limLoc(MathMLElement):
    """
    This element specifies the location of limits in n-ary operators. Limits can be either centered
    above and below the n-ary operator, or positioned just to the right of the operator. When this
    element is omitted, the default location is undOvr.

    Parent elements: naryPr
    """

    pass


class f(MathMLElement):
    """
    This element specifies the fraction object, consisting of a numerator and denominator separated
    by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction
    properties. The fraction object is also used to represent the stack function, which places one
    element above another, with no fraction bar.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: den, fPr, num
    """

    pass


class den(MathMLElement):
    """
    This element specifies the denominator of a fraction.

    Parent element: f

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class num(MathMLElement):
    """
    This element specifies the numerator of a fraction.

    Parent element: f

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class fPr(MathMLElement):
    """
    This element specifies the properties of the fraction function f. Properties of the Fraction function include the
    type or style of the fraction. The fraction bar can be horizontal or diagonal, depending on the fraction
    properties. The fraction object is also used to represent the stack function, which places one element above
    another, with no fraction bar.

    Parent element: f

    Child elements: ctrlPr, type
    """

    pass


class ctrlPr(MathMLElement):
    """
    This element specifies properties on control characters; that is, object characters that cannot
    be selected. Examples of control characters are n-ary operators (excluding their limits and bases),
    fraction bars (excluding the numerator and denominator), and grouping characters (excluding the base).
    ctrlPr allows formatting properties to be stored on these control characters. The control character
    inherits its formatting from the paragraph formatting; ctrlPr contains the formatting differences
    between the control character and the paragraph formatting.

    Parent elements: accPr, barPr, borderBoxPr, boxPr, deg, den, dPr, e, eqArrPr, fName, fPr, funcPr,
    groupChrPr, lim, limLowPr, limUppPr, mPr, naryPr, num, phantPr, radPr, sPrePr, sSubPr, sSubSupPr,
    sSupPr, sub, sup

    Child elements: del, ins, rPr
    """

    pass


class type(MathMLElement):
    """
    This element specifies the type of fraction f; the default is 'bar'. Fractions types are:
    - Stacked Fraction
    - Skewed Fraction
    - Linear Fraction
    - Stack Object (No-Bar Fraction)

    Parent element: fPr
    """

    pass


class func(MathMLElement):
    """
    This element specifies the Function-Apply function, which consists of a function name and an
    argument acted upon.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: e, fName, funcPr
    """

    pass


class fName(MathMLElement):
    """
    This element specifies the name of the function in the Function-Apply object func. For example,
    function names are sin and cos.

    Parent element: func

    Child elements: acc, argPr, bar, bookmarkEnd, bookmarkStart, borderBox, box, commentRangeEnd, commentRangeStart,
    ctrlPr, customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart,
    customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart,
    d, del, eqArr, f, func, groupChr, ins, limLow, limUpp, m, moveFrom, moveFromRangeEnd, moveFromRangeStart, moveTo,
    moveToRangeEnd, moveToRangeStart, nary, oMath, oMathPara, permEnd, permStart, phant, proofErr,
    r, rad, sPre, sSub, sSubSup, sSup
    """

    pass


class funcPr(MathMLElement):
    """
    This element specifies properties such as ctrlPr that can be stored on the function apply object func.

    Parent element: func

    Child element: ctrlPr
    """

    pass


class acc(MathMLElement):
    """
    This element specifies the accent function, consisting of a base and a combining diacritical mark.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: accPr, e
    """

    pass


class accPr(MathMLElement):
    """
    This element specifies the properties of the Accent function.

    Parent element: acc

    Child elements: ctrlPr, chr
    """

    pass


class r(MathMLElement):
    """
    This element specifies a run of math text.

    Parent elements: deg, del, den, e, fName, ins, lim, moveFrom, moveTo, num, oMath, sub, sup

    Child elements: annotationRef, br, commentReference, continuationSeparator, cr, dayLong, dayShort,
    delInstrText, delText, drawing, endnoteRef, endnoteReference, fldChar, footnoteRef, footnoteReference,
    instrText, lastRenderedPageBreak, monthLong, monthShort, noBreakHyphen, object, pgNum, pict, ptab,
    rPr, ruby, separator, softHyphen, sym, t, tab, yearLong, yearShort
    """

    def getTextElements(self):
        # Get all the w:t nodes within the run
        return self.findAll('t')

    def getText(self):
        # Returns a string with all the text in the run
        text = ''
        for t in self.getTextElements():
            text = text + t.text
        return text

    def hasProperty(self, tag, attrib=None):
        mrPr = self.findFirst('rPr')
        if mrPr is None:
            return False
        else:
            if mrPr.hasProperty(tag, attrib):
                return True
            else:
                return False

    def hasStyleProperty(self, s):
        if s == 'nor':
            return self.isRoman()
        if s == 'double-struck':
            return self.isDoubleStruck()
        return None

    def isRoman(self):
        return self.hasProperty('nor')

    def isDoubleStruck(self):
        return self.hasProperty('scr', {'val':'double-struck'})

    def getProperties(self):
        return self.findFirst('rPr')

    def getPreviousContiguousRun(self):
        prevnode = self.getprevious()
        if prevnode is None or prevnode.tag != M+'r':
            return None
        return prevnode

    def getNextContiguousRun(self):
        nextnode = self.getnext()
        if nextnode is None or nextnode.tag != M+'r':
            return None
        return nextnode

    def isFirstWithStyleProperty(self, s):
        prevR = self.getPreviousContiguousRun()
        if prevR is not None:
            return not prevR.hasStyleProperty(s)
        else:
            return True

    def isLastWithStyleProperty(self, s):
        nextR = self.getNextContiguousRun()
        if nextR is not None:
            return not nextR.hasStyleProperty(s)
        else:
            return True


class rPr(MathMLElement):
    """
    This element specifies the properties of the math run r.

    Parent element: r

    Child elements: aln, brk, lit, nor, scr, sty
    """

    def hasProperty(self, tag, attrib=None):

        finds = self.findAll(tag)
        if attrib == None and len(finds)>0:
            return True
        result = False
        for f in finds:
            for k in list(attrib.keys()):
                if k in list(self.keys()):
                    if self.attrib[k] == attrib[k]:
                        result = True
        return result


class scr(MathMLElement):
    """
    This element describes the script applied to the characters in the run. The XML includes the ASCII value of the
    character along with the script of the character. The application maps the ASCII value and script type to the
    appropriate Unicode range.

    Parent element: rPr
    """

    def getScriptValue(self):
        return self.get('val')


class t(MathMLElement):
    """
    This element specifies the text in a math run r.

    Parent elements: r
    """

    def getText(self):
        return self.text


M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = "{%s}" % M_NS
docxBase.NSMAP['m'] = M_NS

namespace = docxBase.lookup.get_namespace(M_NS)
namespace.update(vars())


