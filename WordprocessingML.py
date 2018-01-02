###!/Library/Frameworks/Python.framework/Versions/3.1/bin/python

import docxBase


### WordprocessingML base element class		

class WordprocessingMLElement(docxBase.OfficeOpenXMLElement):
    def _init(self):
        self.nsprefix = 'w'


### WordprocessingML element classes

class document(WordprocessingMLElement):
    """
    This element specifies the contents of a main document part in a WordprocessingML document.
    
    Parent elements: Root element of WordprocessingML Main Document part
    
    Child elements: background, body
    """

    pass


class body(WordprocessingMLElement):
    """
    This element specifies the contents of the body of the document - the main document editing surface.
    The document body contains what is referred to as block-level markup - markup which can exist as a
    sibling element to paragraphs in a WordprocessingML document.
    
    Parent elements: document
    
    Child elements: altChunk, bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml,
    customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd,
    customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, ins, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, p, permEnd, permStart, proofErr,
    sdt, sectPr, tbl
    """

    pass


class WordprocessingMLPropertiesElement(WordprocessingMLElement):
    """
    A base class for the paragraph and run properties elements.
    """

    def addProperties(self, props):
        for k in list(props.keys()):
            y = self.findFirst(k)
            if not y:
                pr = eval(k+'()')
                self.append(pr)
            propAttrs = props[k]
            for attr in list(propAttrs.keys()):
                pr.set(W+attr, propAttrs[attr])

    def hasProperty(self, tag, attrib=None):
        finds = self.findAll(tag)
        if attrib == None and len(finds)>0:
            return True

        result = False
        for f in finds:
            for k in attrib.keys():
                if k in f.attrib.keys():
                    if f.attrib[k] == attrib[k]:
                        result = True
        return result


class pPr(WordprocessingMLPropertiesElement):
    """
    This element specifies a set of paragraph properties which shall be applied to the
    contents of the parent paragraph after all style/numbering/table properties have been
    applied to the text. These properties are defined as direct formatting, since they are
    directly applied to the paragraph and supersede any formatting from styles.
    
    Parent element: p
    
    Child elements: adjustRightInd, autoSpaceDE, autoSpaceDN, bidi, cnfStyle, contextualSpacing,
    divId, framePr, ind, jc, keepLines, keepNext, kinsoku, mirrorIndents, numPr, outlineLvl, overflowPunct,
    pageBreakBefore, pBdr, pPrChange, pStyle, rPr, sectPr, shd, snapToGrid, spacing, suppresAutoHyphens,
    suppressLineNumbers, suppressOverlap, tabs, textAlignment, textboxTightWrap, textDirection, topLinePunct,
    widowControl, wordWrap
    """

    pass


class rPr(WordprocessingMLPropertiesElement):
    """
    This element specifies a set of run properties which shall be applied to the contents of the parent run
    after all style formatting has been applied to the text. These properties are defined as direct formatting,
    since they are directly applied to the run and supersede any formatting from styles.
    
    Parent elements: ctrlPr, r (and pPr for the paragraph mark)
    
    Child elements: b, bCs, bdr, caps, color, cs, dstrike, eastAsianLayout, effect, em, emboss, fitText, highlight,
    i, iCs, imprint, kern, lang, noProof, oMath, outline, position, rFonts, rPrChange, rStyle, rtl, shadow, shd,
    smallCaps, snapToGrid, spacing, specVanish, strik, sz, szCs, u, vanish, vertAlign, w, webHidden
    """

    pass


class p(WordprocessingMLElement):
    """
    This element specifies a paragraph of content in the document.
    The contents of a paragraph in a WordprocessingML document shall consist of any combination of the
    following four types of content: Paragraph properties, Annotations (bookmarks, comments, revisions),
    Custom markup, Run level content (fields, hyperlinks, runs)
    
    Parent elements: body, comment, customXml, docPartBody, endnote, footnote, ftr, hdr, sdtContent, tc, txbxContent
    
    Child elements: customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd,
    customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, fldSimple, hyperlink, ins, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, permEnd, permStart, pPr, proofErr, r,
    sdt, smartTag, subDoc
    """

    def _init(self):
        WordprocessingMLElement._init(self)
        self.contiguous_ignores = [W+'bookmarkStart', W+'bookmarkEnd']

    def getNextParagraph(self):
        nextnode = self.getnext()
        if nextnode is None:
            return None
        while nextnode.tag != W + 'p':
            nextnode = nextnode.getnext()
            if nextnode is None:
                return None
        return nextnode

    def getNextContiguousParagraph(self):
        nextnode = self.getnext()
        while nextnode is not None and nextnode.tag in self.contiguous_ignores:
            nextnode = nextnode.getnext()
        if nextnode is None or nextnode.tag != W+'p':
            return None
        return nextnode

    def getPreviousParagraph(self):
        previousnode = self.getprevious()
        if previousnode is None:
            return None
        while previousnode.tag != W + 'p':
            previousnode = previousnode.getprevious()
            if previousnode is None:
                return None
        return previousnode

    def getPreviousContiguousParagraph(self):
        prevnode = self.getprevious()
        while prevnode is not None and prevnode.tag in self.contiguous_ignores:
            prevnode = prevnode.getprevious()
        if prevnode is None or prevnode.tag != W+'p':
            return None
        return prevnode

    def getRuns(self):
        return self.findAll('r')

    def getTextElements(self):
        return self.findAll('t') #doXpath('w:r/w:t')

    def getText(self):
        text = ''
        for t in self.getTextElements():
            text = text + t.text
        return text

    def hasText(self):
        return len(self.getText())>0

    def preserveSpace(self):
        for t in self.getTextElements():
            t.preserveSpace()

    def hasProperty(self, tag, attrib=None):
        pPr = self.findFirst('pPr')
        if pPr is None:
            return False
        else:
            if pPr.hasProperty(tag, attrib):
                return True
            else:
                return False

    def addProperties(self, props):
        wpPr = self.findFirst('pPr')
        if not wpPr:
            wpPr = pPr()
            self.insert(0, wpPr)
        wpPr.addProperties(props)
        return wpPr

    def removeProperty(self, prop):
        x = self.doXpath('w:pPr/w:'+prop)
        if not x:
            return
        else:
            self.findAll('pPr')[0].remove(x[0])

    def addRunProperties(self, props):
        runs = self.findAll('r')
        if not runs:
            return
        for r in runs:
            r.addProperties(props)

    def getStyle(self):
        pst = self.findAll('pStyle')
        if len(pst)>0:
            return pst[0].get('val')
        else:
            return None

    def hasStyle(self, s):
        if self.getStyle() == s:
            return True
        return False

    def isFirstWithStyle(self):
        s = self.getStyle()
        if s is None:
            return False
        prevp = self.getPreviousContiguousParagraph()
        if prevp is not None and prevp.hasStyle(s):
            return False
        return True

    def hasBookmark(self):
        if len(self.findAll('bookmarkStart'))>0:
            return True
        else:
            return False

    def hasNumbering(self):
        if self.hasProperty('numPr'):
            np = self.findFirst('numPr')
            ni = np.findFirst('numId')
            il = np.findFirst('ilvl')
            if ni is None or il is None:
                return False
            else:
                return True
        return False

    def getNumbering(self):
        if self.hasNumbering():
            numPr = self.findFirst('numPr')
            numId = numPr.findFirst('numId')
            ilvl = numPr.findFirst('ilvl')
            numIdVal = numId.get('val')
            ilvlVal = ilvl.get('val')
            return [numIdVal, ilvlVal]
        return None

    def hasCaptionField(self, floatnames):
        flds = self.findAll('fldSimple')
        for f in flds:
            fcode = f.get('instr').split()
            if fcode[0] == 'SEQ' and fcode[1] in floatnames:
                return True
        flds = self.findAll('instrText')
        for f in flds:
            fcode = f.text.split()
            if fcode[0] == 'SEQ' and fcode[1] in floatnames:
                return True
        return False

    def isFootnote(self):
        for n in self.iterancestors():
            if n.tag == W+'footnote':
                return True
        return False

    def isListItem(self, numprops=None):
        isListItem = False
        if self.hasNumbering():
            this_numId, this_ilvl = self.getNumbering()
            if numprops is not None:
                numId, ilvl = numprops
                if this_numId == numId and this_ilvl == ilvl:
                    isListItem = True
            else:
                isListItem = True
        return isListItem

    def isFirstListItem(self):
        prevP = self.getPreviousContiguousParagraph()
        while 1:
            if prevP is None or prevP.hasText():
                break
            prevP = prevP.getPreviousContiguousParagraph()

        if prevP is not None:
            return not prevP.isListItem(self.getNumbering())
        else:
            return True

    def isLastListItem(self):
        nextP = self.getNextContiguousParagraph()
        while 1:
            if nextP is None or nextP.hasText():
                break
            nextP = nextP.getNextContiguousParagraph()

        if nextP is not None:
            return not nextP.isListItem(self.getNumbering())
        else:
            return True

    def isHeading(self, hstyles):
        ps = self.getStyle()
        if ps in list(hstyles.keys()):
            return hstyles[ps]
        else:
            return None

    def isCaption(self, cstyles):
        ps = self.getStyle()
        if ps in list(cstyles.keys()):
            return True
        return False

    def isTableNote(self, tnstyles):
        ps = self.getStyle()
        if ps in list(tnstyles): #.keys()):
            return True
        return False

    def isEnvironment(self, style2environment):
        pstyle = self.getStyle()
        if pstyle is not None:
            if pstyle.lower() in style2environment.keys():
                return True
        return False

    def getEnvironment(self, style2environment):
        pstyle = self.getStyle()
        environment = None
        if pstyle is not None:
            if pstyle.lower() in style2environment.keys():
                environment = style2environment[pstyle.lower()]
        return environment

    def isFirstWithEnvironment(self):
        s = self.getStyle()
        prevP = self.getPreviousContiguousParagraph()
        #while 1:
        #	if prevP is None or prevP.hasText():
        #		break
        #	prevP = prevP.getPreviousContiguousParagraph()

        if prevP is not None and prevP.getStyle() == s:
            return False
        return True

    def isLastWithEnvironment(self):
        s = self.getStyle()
        nextP = self.getNextContiguousParagraph()
        #while 1:
        #	if nextP is None or nextP.hasText():
        #		break
        #	nextP = nextP.getNextContiguousParagraph()

        if nextP is not None and nextP.getStyle() == s:
            return False
        return True


class r(WordprocessingMLElement):
    """
    This element specifies a run of content in the parent field, hyperlink, custom XML element, structured document tag,
    smart tag, or paragraph. The contents of a run in a WordprocessingML document shall consist of any combination of run content.
    
    Parent elements: customXml, del, fldSimple, hyperlink, ins, moveFrom, moveTo, p, rt, rubyBase, sdtContent, smartTag
    
    Child elements: annotationRef, br, commentReference, continuationSeparator, cr, dayLong, dayShort, delInstrText, delText, drawing,
    endnoteRef, endnoteReference, fldChar, footnoteRef, footnoteReference, instrText, lastRenderedPageBreak, monthLong, monthShort,
    noBreakHyphen, object, pgNum, pict, ptab, rPr, ruby, separator, softHyphen, sym, t, tab, yearLong, yearShort
    """

    def _init(self):
        WordprocessingMLElement._init(self)
        self.contiguous_ignores = [W+'bookmarkStart', W+'bookmarkEnd']

    def getPreviousContiguousRun(self):
        prevnode = self.getprevious()
        while prevnode is not None and prevnode.tag in self.contiguous_ignores:
            prevnode = prevnode.getprevious()
        if prevnode is None or prevnode.tag != W+'r':
            return None
        return prevnode

    def getNextContiguousRun(self):
        nextnode = self.getnext()
        while nextnode is not None and nextnode.tag in self.contiguous_ignores:
            nextnode = nextnode.getnext()
        if nextnode is None or nextnode.tag != W+'r':
            return None
        return nextnode

    def getTextElements(self):
        # Get all the w:t nodes within the run
        return self.findAll('t')

    def hasText(self):
        if len(self.getTextElements())>0:
            return True
        else:
            return False

    def getText(self):
        # Returns a string with all the text in the run
        text = ''
        for t in self.getTextElements():
            text = text + t.text
        return text

    def isBold(self, useDirect=True):
        if useDirect:
            return self.isBoldDirect()
        return (self.getStyle() == 'ZinkBold') or (self.getStyle() == 'ZinkBoldItalic') or (self.getStyle() == 'ZinkBoldUnderline') or (self.getStyle() == 'ZinkBoldItalicUnderline')

    def isBoldDirect(self):
        return self.hasProperty('b') and not self.hasProperty('b', {W+'val':'false'})

    def isItalic(self, useDirect=True):
        if useDirect:
            return self.isItalicDirect()
        return (self.getStyle() == 'ZinkItalic') or (self.getStyle() == 'ZinkBoldItalic') or (self.getStyle() == 'ZinkItalicUnderline') or (self.getStyle() == 'ZinkBoldItalicUnderline')

    def isItalicDirect(self):
        return self.hasProperty('i') and not self.hasProperty('i', {W+'val':'false'})

    def isUnderline(self, useDirect=True):
        if useDirect:
            return self.isUnderlineDirect()
        return (self.getStyle() == 'ZinkUnderline') or (self.getStyle() == 'ZinkItalicUnderline') or (self.getStyle() == 'ZinkBoldUnderline') or (self.getStyle() == 'ZinkBoldItalicUnderline')

    def isUnderlineDirect(self):
        return self.hasProperty('u') and not self.hasProperty('u', {W+'val':'false'})

    def isSubScript(self):
        return self.hasProperty('vertAlign', {W+'val':'subscript'})

    def isSuperScript(self):
        return self.hasProperty('vertAlign', {W+'val':'superscript'})

    def getStyle(self):
        rprop = self.getProperties()
        if rprop is not None:
            rs = rprop.findFirst('rStyle')
            if rs is not None:
                return rs.get('val')

    def hasStyleProperty(self, s, useDirect=True):
        if s == 'b':
            return self.isBold(useDirect)
        if s == 'i':
            return self.isItalic(useDirect)
        if s == 'u':
            return self.isUnderline(useDirect)
        if s == 'superscript':
            return self.isSuperScript()
        if s == 'subscript':
            return self.isSubScript()
        return None

    def styledNeighborsHaveText(self, styleProp, useDirect=True):
        text = self.getText()

        # Forward
        next = self.getNextContiguousRun()
        while next is not None and next.hasStyleProperty(styleProp, useDirect):
            text = text + next.getText()
            next = next.getNextContiguousRun()

        # Backward
        prev = self.getPreviousContiguousRun()
        while prev is not None and prev.hasStyleProperty(styleProp, useDirect):
            text = text + prev.getText()
            prev = prev.getPreviousContiguousRun()

        if text.strip() == '':
            return False
        else:
            return True

    def hasProperty(self, tag, attrib=None):
        wrPr = self.findFirst('rPr')
        if wrPr is None:
            return False
        else:
            if wrPr.hasProperty(tag, attrib):
                return True
            else:
                return False

    def getProperties(self):
        return self.findFirst('rPr')

    def addProperties(self, props):
        wrPr = self.findFirst('rPr')
        if wrPr is None:
            wrPr = rPr()
            self.insert(0, wrPr)
        wrPr.addProperties(props)
        return wrPr

    def removeProperty(self, prop):
        x = self.doXpath('w:rPr/w:'+prop)
        if not x:
            return
        else:
            self.findFirst('rPr').remove(x[0])

    def hasFootnote(self):
        if self.findAll('footnoteReference'):
            return True
        else:
            return False

    def getFootnoteId(self):
        if self.hasFootnote():
            fnr = self.findFirst('footnoteReference')
            return fnr.get('id')
        else:
            return None

    def getLanguage(self):
        if self.hasProperty('lang'):
            return self.findFirst('rPr').findFirst('lang').get('val')
        else:
            return None

    def getFont(self):
        # rFonts could have more than one unique font for different char ranges!! should be incorporated
        if self.hasProperty('rFonts'):
            return self.findFirst('rPr').findFirst('rFonts').get('ascii')
        else:
            return None

    def isFirstWithStyleProperty(self, s, useDirect=True):
        if not self.styledNeighborsHaveText(s, useDirect):
            return False

        prevR = self.getPreviousContiguousRun()
        if prevR is not None:
            return not prevR.hasStyleProperty(s, useDirect)
        else:
            return True

    def isLastWithStyleProperty(self, s, useDirect=True):
        if not self.styledNeighborsHaveText(s, useDirect):
            return False

        nextR = self.getNextContiguousRun()
        if nextR is not None:
            return not nextR.hasStyleProperty(s, useDirect)
        else:
            return True


class t(WordprocessingMLElement):
    """
    This element specifies that this run contains literal text which shall be displayed in the document.
    The t element shall be used for all text runs which is not: Part of a region of text that is contained
    in a deleted region using the del element; Part of a region of text that is contained within a field code
    
    Parent element: r
    """

    def checkSpace(self):
        # check whether text is present, and whether it has preceding or trailing spaces
        if self.text and self.text != self.text.strip():
            self.ensurePreserveSpace()
        else:
            self.ensureNoPreserveSpace()

    def ensurePreserveSpace(self):
        # Check for the existence of the space='preserve' attribute
        if 'space' in self.attrib:
            if self.attrib[W+'space'] == 'preserve':
                return
            else:
                self.attrib[W+'space'] = 'preserve'
        else:
            self.set(W+'space', 'preserve')

    def ensureNoPreserveSpace(self):
        # remove the space attribute if present
        if 'space' in self.attrib:
            del self.attrib[W+'space']

    def getText(self):
        return self.text


### Footnote element classes


class footnotes(WordprocessingMLElement):
    """
    This element specifies the set of all footnotes in the document, including footnote separators and
    continuation notices. This element is the root node for the Footnotes part.
    
    Parent elements: Root element of WordprocessingML Footnotes part
    
    Child elements: footnote
    """

    pass


class footnote(WordprocessingMLElement):
    """
    This element specifies the content of a single footnote within a WordprocessingML document. Each footnote
    shall be represented by a single footnote element, which may contain any valid block-level content.
    
    Parent elements: footnotes
    
    Child elements: altChunk, bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml,
    customXmlDelRangeEnd, customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd,
    customXmlMoveFromRangeStart, customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, ins, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, p, permEnd, permStart, proofErr,
    sdt, tbl
    """

    def getParagraphs(self):
        ps = self.findAll('p')
        return ps


class footnoteReference(WordprocessingMLElement):
    """
    This element specifies the presence of a footnote reference. A footnote reference is a run of automatically
    numbered text which references a particular footnote within the parent document, and inherits the footnote
    reference mark's numbering.
    If an footnote reference is specified within a footnote or endnote, then the document shall be considered
    non-conformant.
    
    Parent element: r
    """

    def getFootnoteId(self):
        return self.get('id')


class footnoteRef(WordprocessingMLElement):
    """
    This element specifies the presence of a footnote reference mark. A footnote reference mark is a run of
    automatically numbered text which follows the numbering format set forth via the footnote numFmt element.
    If a footnote reference mark is specified within a run which is not part of a footnote, then that footnote
    reference mark may be ignored.
    
    Parent element: r
    """

    pass


class separator(WordprocessingMLElement):
    """
    This element specifies the presence of a separator mark within the current run. A separator mark is a horizontal
    line which spans part of the width text extents.
    Note: The separator mark is typically used within the context of separator footnotes or endnotes. These
    footnote and endnote types define the footnote/endnote used to separate the contents of the main document
    story from the contents of footnotes or endnotes on that page.
    
    Parent element: r
    """

    pass


class continuationSeparator(WordprocessingMLElement):
    """
    This element specifies the presence of a continuation separator mark within the current run. A continuation
    separator mark is a horizontal line which spans the width of the main story's text extents.
    Note: The continuation separator mark is typically used within the context of continuation separator footnotes
    or endnotes. These footnote and endnote types define the footnote/endnote used to separate the contents of
    the main document story from continuation of footnotes or endnotes which began on a previous page.
    
    Parent element: r
    """

    pass


class footnotePr(WordprocessingMLElement):
    """
    This element specifies the footnote properties for this document or for the current section. Each property is stored as a unique element
    within the footnotePr element.
    Document-wide footnote properties may be overridden for a specific section via the section-wide footnotePr element.
    If this element is omitted for a given section, then that section shall use the footnote properties defined at the document-wide level.
    
    Parent elements: sectPr, settings
    
    Child elements: footnote, numFmt, numRestart, numStart, pos
    """

    pass


class numRestart(WordprocessingMLElement):
    """
    This element specifies when all automatic numbering for the footnote or endnote reference marks shall be
    restarted. When restarted, the next automatically numbered footnote or endnote in the document (each type
    handled independently) shall restart to the specified numStart value.
    If this element is omitted, then automatic numbering shall not be restarted between each page or section (a
    value of continuous).
    
    Parent elements: endnotePr, footnotePr
    """

    pass





### Miscellaneous element classes	

### Docx property element classes

### Run property classes


class i(WordprocessingMLElement):
    """
    This element specifies whether the italic property should be applied to all non-complex script characters in the
    contents of this run when displayed in a document.
    This formatting property is a toggle property, which specifies that its behavior differs between its use within a
    style definition and its use as direct formatting. When used as part of a style definition, setting this property
    shall toggle the current state of that property as specified up to this point in the hierarchy (i.e. applied to not
    applied, and vice versa). Setting it to false (or an equivalent) shall result in the current setting remaining
    unchanged. However, when used as direct formatting, setting this property to true or false shall set the
    absolute state of the resulting property.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy .If this element is never applied in the style hierarchy, then italics shall not be applied to non-complex
    script characters.
    
    Parent element: rPr
    """

    pass


class b(WordprocessingMLElement):
    """
    This element specifies whether the bold property shall be applied to all non-complex script characters
    in the contents of this run when displayed in a document.
    This formatting property is a toggle property, which specifies that its behavior differs between its use
    within a style definition and its use as direct formatting. When used as part of a style definition, setting
    this property shall toggle the current state of that property as specified up to this point in the hierarchy
    (i.e. applied to not applied, and vice versa). Setting it to false (or an equivalent) shall result in the
    current setting remaining unchanged. However, when used as direct formatting, setting this property to true
    or false shall set the absolute state of the resulting property.
    If this element is not present, the default value is to leave the formatting applied at previous level in
    the style hierarchy. If this element is never applied in the style hierarchy, then bold shall not be applied
    to non-complex script characters.
    
    Parent element: rPr
    """

    pass


class u(WordprocessingMLElement):
    """
    This element specifies that the contents of this run should be displayed along with an underline appearing
    directly below the character height (less all spacing above and below the characters on the line).
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then an underline shall not be applied to the
    contents of this run.
    
    Parent element: rPr
    """

    patterns = ('single', 'words', 'double', 'thick', 'dotted', 'dottedHeavy', 'dash', 'dashedHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dashDotHeavy', 'dotDotDash', 'dashDotDotHeavy', 'wave', 'wavyHeavy', 'wavyDouble', 'none')


class iCs(WordprocessingMLElement):
    """
    This element specifies whether the italic property should be applied to all complex script characters in the
    contents of this run when displayed in a document.
    
    Parent element: rPr
    """

    pass


class bCs(WordprocessingMLElement):
    """
    This element specifies whether the bold property shall be applied to all complex script characters in the contents
    of this run when displayed in a document.
    
    Parent element: rPr
    """

    pass


class smallCaps(WordprocessingMLElement):
    """
    This element specifies that all small letter characters in this text run shall be formatted for display only as their
    capital letter character equivalents in a font size two points smaller than the actual font size specified for this
    text. This property does not affect any non-alphabetic character in this run, and does not change the Unicode
    character for lowercase text, only the method in which it is displayed. If this font cannot be made two point
    smaller than the current size, then it shall be displayed as the smallest possible font size in capital letters.
    This formatting property is a toggle property, which specifies that its behavior differs between its use within a
    style definition and its use as direct formatting. When used as part of a style definition, setting this property
    shall toggle the current state of that property as specified up to this point in the hierarchy (i.e. applied to not
    applied, and vice versa). Setting it to false (or an equivalent) shall result in the current setting remaining
    unchanged. However, when used as direct formatting, setting this property to true or false shall set the
    absolute state of the resulting property.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then the characters are not formatted as capital
    letters.
    This element shall not be present with the caps property on the same run, since they are mutually exclusive in terms of
    appearance.
    
    Parent element: rPr
    """

    pass


class highlight(WordprocessingMLElement):
    """
    This element specifies a highlighting color which is applied as a background behind the contents of this run.
    If this run has any background shading specified using the shd element, then the background shading
    shall be superseded by the highlighting color when the contents of this run are displayed.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then text highlighting shall not be applied to the
    contents of this run.
    
    Parent element: rPr
    """

    pass


class rFonts(WordprocessingMLElement):
    """
    This element specifies the fonts which shall be used to display the text contents of this run. Within a single run,
    there may be up to four types of content present which shall each be allowed to use a unique font:
    ASCII; High ANSI; Complex Script; East Asian
    The use of each of these fonts shall be determined by the Unicode character values of the run content,
    unless manually overridden via use of the cs element.
    If this element is not present, the default value is to leave the formatting applied at previous level in the
    style hierarchy. If this element is never applied in the style hierarchy, then the text shall be displayed in any
    default font which supports each type of content.
    
    Parent element: rPr
    """

    pass


class sz(WordprocessingMLElement):
    """
    This element specifies the font size which shall be applied to all non complex script characters in the contents of
    this run when displayed.
    If this element is not present, the default value is to leave the value applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then any appropriate font size may be used for
    non complex script characters.
    Value is specified in half-points
    
    Parent element: rPr
    """

    pass


class szCs(WordprocessingMLElement):
    """
    This element specifies the font size which shall be applied to all complex script characters in the contents of this
    run when displayed.
    If this element is not present, the default value is to leave the value applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then any appropriate font size may be used for
    complex script characters.
    Value is specified in half-points
    
    Parent element: rPr
    """

    pass


class lang(WordprocessingMLElement):
    """
    This element specifies the languages which shall be used to check spelling and grammar (if requested) when
    processing the contents of this run.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then the languages for the contents of this run
    shall be automatically determined based on their contents using any method desired.
    
    Parent element: rPr
    """

    pass


class rStyle(WordprocessingMLElement):
    """
    This element specifies the style ID of the character style which shall be used to format the contents of this
    paragraph.
    This formatting is applied at the following location in the style hierarchy:
    Document defaults; Table styles; Numbering styles; Paragraph styles; Character styles (this element); Direct Formatting
    This means that all properties specified in the style element with a styleId which corresponds to the value in this
    element's val attribute are applied to the run at the appropriate level in the hierarchy.
    If this element is omitted, or it references a style which does not exist, then no character style shall be applied to
    the current paragraph. As well, this property is ignored if the run properties are part of a character style.
    
    Parent element: rPr
    """

    pass


class color(WordprocessingMLElement):
    """
    This element specifies the color which shall be used to display the contents of this run in the document.
    This color may be explicitly specified, or set to allow the consumer to automatically choose an appropriate color
    based on the background color behind the run's content.
    
    Parent element: rPr
    """

    pass


class vertAlign(WordprocessingMLElement):
    """
    This element specifies the alignment which shall be applied to the contents of this run in relation to the default
    appearance of the run's text. This allows the text to be repositioned as subscript or superscript without altering
    the font size of the run properties.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then the text shall not be subscript or
    superscript relative to the default baseline location for the contents of this run.
    
    Parent element: rPr
    """

    def isSuperScript(self):
        if self.attrib[W+'val'] == 'superscript':
            return True
        return False

    def isSubScript(self):
        if self.attrib[W+'val'] == 'subscript':
            return True
        return False

    def getValue(self):
        return self.attrib[W+'val']

### Paragraph property classes	


class pStyle(WordprocessingMLElement):
    """
    This element specifies the style ID of the paragraph style which shall be used to format the contents of this
    paragraph.
    This formatting is applied at the following location in the style hierarchy:
    Document defaults; Table styles; Numbering styles; Paragraph styles (this element); Character styles; Direct Formatting
    This means that all properties specified in the style element with a styleId which corresponds to the
    value in this element's val attribute are applied to the paragraph at the appropriate level in the hierarchy.
    If this element is omitted, or it references a style which does not exist, then no paragraph style shall be applied to
    the current paragraph. As well, this property is ignored if the paragraph properties are part of a paragraph style.
    
    Parent element: pPr
    """

    pass


class ind(WordprocessingMLElement):
    """
    This element specifies the set of indentation properties applied to the current paragraph.
    Indentation settings are overriden on an individual basis - if any single attribute on this element is omitted on a
    given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e. that
    previous setting remains unchanged). If any single attribute on this element is never specified in the style
    hierarchy, then no indentation of that type is applied to the paragraph.
    
    Parent element: pPr
    """

    pass


class spacing(WordprocessingMLElement):
    """
    This element specifies the inter-line and inter-paragraph spacing which shall be applied to the contents of this
    paragraph when it is displayed by a consumer.
    If this element is omitted on a given paragraph, each of its values is determined by the setting previously set at
    any level of the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in
    the style hierarchy, then the paragraph shall have no spacing applied to its lines, or above and below its contents.
    
    Parent element: pPr
    """

    pass


class jc(WordprocessingMLElement):
    """
    This element specifies the paragraph alignment which shall be applied to text in this paragraph.
    If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level
    of the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in the style
    hierarchy, then no alignment is applied to the paragraph.
    
    Parent element: pPr
    """

    pass


class autoSpaceDE(WordprocessingMLElement):
    """
    This element specifies whether inter-character spacing shall automatically be adjusted between regions of Latin text and
    regions of East Asian text in the current paragraph. These regions shall be determined by the Unicode character values
    of the text content within the paragraph.
    Note: This property is used to ensure that the spacing between regions of Latin text and adjoining East Asian text is
    sufficient on each side such that the Latin text can be easily read within the East Asian text.
    If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of
    the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in the style
    hierarchy, its value is assumed to be true.
    
    Parent element: pPr
    """

    pass


class autoSpaceDN(WordprocessingMLElement):
    """
    This element specifies whether inter-character spacing shall automatically be adjusted between regions of numbers and
    regions of East Asian text in the current paragraph. These regions shall be determined by the Unicode character values
    of the text content within the paragraph.
    Note: This property is used to ensure that the spacing between regions of numbers and adjoining East Asian text is
    sufficient on each side such that the numbers can be easily read within the East Asian text.
    If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of
    the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in the style
    hierarchy, its value is assumed to be true.
    
    Parent element: pPr
    """

    pass


class adjustRightInd(WordprocessingMLElement):
    """
    This element specifies whether the right indent shall be automatically adjusted for the given paragraph when a
    document grid has been defined for the current section using the docGrid element, modifying of the current right
    indent used on this paragraph.
    Note: This setting is used in order to ensure that the line breaking for that paragraph is not determined by the
    width of the final character on the line.
    If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level
    of the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in the
    style hierarchy, its value is assumed to be true.
    
    Parent element: pPr
    """

    pass


class framePr(WordprocessingMLElement):
    """
    This element specifies information about the current paragraph with regard to text frames. Text
    frames are paragraphs of text in a document which are positioned in a separate region or frame
    in the document, and can be positioned with a specific size and position relative to non-frame
    paragraphs in the current document.
    The first piece of information specified by the framePr element is that the current paragraph
    is actually part of a text frame in the document. This information is specified simply by the
    presence of the framePr element in paragraph's properties. If the framePr element is omitted,
    the paragraph shall not be part of any text frame in the document.
    The second piece of information concerns the set of paragraphs which are part of the current
    text frame in the document. This is determined based on the attributes on the framePr element.
    If the set of attribute values specified on two adjacent paragraphs is identical, then those
    two paragraphs shall be considered to be part of the same text frame and rendered within the
    same frame in the document.
    The positioning of the frame relative to the properties stored on its attribute values shall be
    calculated relative to the next paragraphs in the document which is itself not part of a text frame.
    
    Parent element: pPr
    """

    pass

### Paragraph properties - Tabs


class tabs(WordprocessingMLElement):
    """
    This element specifies a sequence of custom tab stops which shall be used for any tab characters in the current paragraph.
    If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of
    the style hierarchy (i.e. that previous setting remains unchanged). If this setting is never specified in the style
    hierarchy, then no custom tab stops shall be used for this paragraph.
    As well, this property is additive - tab stops at each level in the style hierarchy are added to each other to determine
    the full set of tab stops for the paragraph. A hanging indent specified via the hanging attribute on the ind element
    shall also always implicitly create a custom tab stop at its location.
    
    Parent element: pPr
    
    Child element: tab
    """

    pass


class tab(WordprocessingMLElement):
    """
    This element specifies a single custom tab stop within a set of custom tab stops applied as part of a set of
    customized paragraph properties in a document.
    
    Parent element: tabs
    """

    pass


### Paragraph Properties - Numbering Properties


class numPr(WordprocessingMLElement):
    """
    This element specifies that the current paragraph references a numbering definition instance in the current
    document.
    The presence of this element specifies that the paragraph will inherit the properties specified by the numbering
    definition in the num element at the level specified by the level specified in the lvl element
    and shall have an associated number positioned before the beginning of the text flow in this paragraph. When
    this element appears as part of the paragraph formatting for a paragraph style, then any numbering level
    defined using the ilvl element shall be ignored, and the pStyle element on the associated abstract
    numbering definition shall be used instead.
    
    Parent element: pPr
    
    Child elements: ilvl, ins, numberingChange, numId
    """

    pass


class ilvl(WordprocessingMLElement):
    """
    This element specifies the numbering level of the numbering definition instance which shall be applied to the
    parent paragraph.
    This numbering level is specified on either the abstract numbering definition's lvl element, and may be
    overridden by a numbering definition instance level override's lvl element.
    
    Parent element: numPr
    TAG  = 'ilvl'
    """

    pass


class numId(WordprocessingMLElement):
    """
    This element specifies the numbering definition instance which shall be used for the given parent numbered
    paragraph in the WordprocessingML document.
    A value of 0 for the val attribute shall never be used to point to a numbering definition instance, and shall
    instead only be used to designate the removal of numbering properties at a particular level in the style hierarchy
    (typically via direct formatting).
    
    Parent element: numPr
    """

    pass


### Paragraph properties - Section Properties


class sectPr(WordprocessingMLElement):
    """
    This element defines the section properties for the a section of the document. Note: For the last section in the
    document, the section properties are stored as a child element of the body element.
    
    Parent element: pPr (and body for the last section in the document)
    
    Child elements: bidi, cols, docgrid, endnotePr, footerReference, footnotePr, formProt, headerReference, lnNumType,
    noEndnote, paperScr, pgBorders, pgMar, pgSz, printerSettings, rtlGutter, sectPrChange, textDirection, titlePg, type,
    vAlign
    """

    pass


class headerReference(WordprocessingMLElement):
    """
    This element specifies a single header which shall be associated with the current section in the document. This
    header shall be referenced via the id attribute, which specifies an explicit relationship to the appropriate Header
    part in the WordprocessingML package.
    If the relationship type of the relationship specified by this element is not
    http://schemas.openxmlformats.org/officeDocument/2006/header, is not present, or does not have a
    TargetMode attribute value of Internal, then the document shall be considered non-conformant.
    Within each section of a document there may be up to three different types of headers:
    - First page header
    - Odd page header
    - Even page header
    The header type specified by the current headerReference is specified via the type attribute. If any type of
    header is omitted for a given section, then the following rules shall apply.
    - If no headerReference for the first page header is specified and the titlePg element is specified, then
      the first page header shall be inherited from the previous section or, if this is the first section in
      the document, a new blank header shall be created. If the titlePg element is not specified, then no
      first page header shall be shown, and the odd page header shall be used in its place.
    - If no headerReference for the even page header is specified and the evenAndOddHeaders element is specified,
      then the even page header shall be inherited from the previous section or, if this is the first section in
      the document, a new blank header shall be created. If the evenAndOddHeaders element is not specified, then
      no even page header shall be shown, and the odd page header shall be used in its place.
    - If no headerReference for the odd page header is specified then the even page header shall be inherited from
      the previous section or, if this is the first section in the document, a new blank header shall be created.
    
    Parent element: sectPr
    """

    pass


class footerReference(WordprocessingMLElement):
    """
    This element specifies a single footer which shall be associated with the current section in the document. This
    footer shall be referenced via the id attribute, which specifies an explicit relationship to the appropriate Footer
    part in the WordprocessingML package.
    If the relationship type of the relationship specified by this element is not
    http://schemas.openxmlformats.org/officeDocument/2006/footer, is not present, or does not have a
    TargetMode attribute value of Internal, then the document shall be considered non-conformant.
    Within each section of a document there may be up to three different types of footers:
    - First page footer
    - Odd page footer
    - Even page footer.
    The footer type specified by the current footerReference is specified via the type attribute.
    
    If any type of footer is omitted for a given section, then the following rules shall apply.
    - If no footerReference for the first page footer is specified and the titlePg element is specified, then the
      first page footer shall be inherited from the previous section or, if this is the first section in the
      document, a new blank footer shall be created. If the titlePg element is not specified, then no first page
      footer shall be shown, and the odd page footer shall be used in its place.
    - If no footerReference for the even page footer is specified and the evenAndOddHeaders element is
      specified, then the even page footer shall be inherited from the previous section or, if this is the first
      section in the document, a new blank footer shall be created. If the evenAndOddHeaders element is not specified,
      then no even page footer shall be shown. and the odd page footer shall be used in its place.
    - If no footerReference for the odd page footer is specified then the even page footer shall be inherited from
      the previous section or, if this is the first section in the document, a new blank footer shall be created.
    
    Parent element: sectPr
    """

    pass


class pgSz(WordprocessingMLElement):
    """
    This element specifies the properties (size and orientation) for all pages in the current section. The size values
    are specified in twentieths of a point.
    
    Parent element: sectPr
    """

    pass


class pgMar(WordprocessingMLElement):
    """
    This element specifies the page margins for all pages in this section. Values are specified in twentieths of a point.
    
    Parent element: sectPr
    """

    pass


class docGrid(WordprocessingMLElement):
    """
    This element specifies the settings for the document grid, which enables precise layout of full-width East Asian
    language characters within a document by specifying the desired number of characters per line and lines per
    page for all East Asian text content in this section.
    
    Parent element: sectPr
    """

    pass


class pgNumType(WordprocessingMLElement):
    """
    This element specifies the page numbering settings for all page numbers that appear in the contents of the
    current section.
    
    Parent element: sectPr
    """

    pass


class cols(WordprocessingMLElement):
    """
    This element specifies the set of columns defined for this section in the document.
    
    Parent element: sectPr
    
    Child element: col
    """

    pass


class titlePg(WordprocessingMLElement):
    """
    This element specifies whether the parent section in this document shall have a different header and footer for
    its first page.
    If the val attribute is set to true, then the parent section in the document shall use a first page header for the
    first page in the section. If the val attribute is set to false, then the first page in the parent section shall use the
    odd page header.
    This setting does not affect the presence of even and odd page header on all sections, which is specified using
    the evenAndOddHeaders element.
    If this element is set to false and a first page header is specified , then it shall be ignored and only the odd page
    header shall be displayed. Conversely, if this element is set to true and the first page header type is omitted for
    the given section, then a blank header shall be created as needed (another header type shall not be used in its
    place).
    If this element is omitted, then its value shall be assumed to be false.
    
    Parent element: sectPr
    """

    pass


class type(WordprocessingMLElement):
    """
    This element specifies the type of the current section. The section type specifies how the contents of the current
    section shall be placed relative to the previous section.
    WordprocessingML supports five distinct types of section breaks:
    - Next page section breaks (the default if type is not specified), which begin the new section on the following page.
    - Odd page section breaks, which begin the new section on the next odd-numbered page.
    - Even page section breaks, which begin the new section on the next even-numbered page.
    - Continuous section breaks, which begin the new section on the following paragraph. This means that
      continuous section breaks might not specify certain page-level section properties, since they must be
      inherited from the following section. These breaks, however, can specify other section properties, such
      as line numbering and footnote/endnote settings.
    - Column section breaks, which begin the new section on the next column on the page.
    
    Parent element: sectPr
    """

    pass


class vAlign(WordprocessingMLElement):
    """
    This element specifies the vertical alignment for text on pages in the current section, relative to the top and
    bottom margins in the main document story on each page.
    
    Parent element: sectPr
    """

    pass

### Table classes

class tbl(WordprocessingMLElement):
    """
    This element specifies the contents of a table present in the document. A table is a set of paragraphs (and other
    block-level content) arranged in rows and columns. Tables in WordprocessingML are defined via the tbl element,
    which is analogous to the HTML table tag.
    
    Parent elements: body, comment, customXml, docPartBody, endnote, footnote, ftr, hdr, sdtContent, tc, txbxContent
    
    Child elements: bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml, customXmlDelRangeEnd,
    customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart,
    customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, ind, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, permEnd, permStart, proofErr, sdt, tblGrid, tblPr, tr
    """

    def getWidth(self):
        tpr = self.findFirst('tblPr')
        if tpr is None:
            return None
        twidth = tpr.findFirst('tblW')
        if twidth is None:
            return None
        return twidth.getValue()

    def getNumberOfLogicalColumns(self):
        return len(self.findFirst('tblGrid').findAll('gridCol'))

    def getMaximumNumberOfRowCells(self):
        maxn = 0
        for r in self.findAllInChildren('tr'):
            ncells = r.getNumberOfCells()
            if ncells > maxn:
                maxn = ncells
        return maxn

    def getNumberOfRows(self):
        return len(self.getRows())

    def getRows(self):
        return self.findAllInChildren('tr')

    def getRow(self, n):
        rows = self.getRows()
        if n >= len(rows):
            raise ValueError('Row index exceeds number of rows')
        return rows[n]

    def getCell(self, i, j):
        return self.getRow(i).getCell(j)

    def getShadingDict(self):
        tPr = self.findFirst('tblPr')
        if tPr is not None:
            tShading = tPr.findFirst('shd')
            if tShading is not None:
                val = tShading.get('val')
                color = tShading.get('color')
                fill = tShading.get('fill')

                return {'val':  val, 'color': color, 'fill': fill}
        return None

    def getBorderDict(self, side):
        tablePr = self.findFirst('tblPr')
        if tablePr is not None:
            tableBorders = tablePr.findFirst('tblBorders')
            if tableBorders is not None:
                b = tableBorders.findFirst(side)
                if b is not None:
                    val = b.get('val')
                    sz = b.get('sz')
                    space = b.get('space')
                    color = b.get('color')

                    return {'val': val, 'sz': sz, 'space': space, 'color': color}
        return None

    def getBorder(self, side):
        bdict = self.getBorderDict(side)
        if bdict is not None:
            return bdict['val']
        else:
            return 'nil'

    def getColumnWidth(self, columnnumber):
        grid = self.findFirst('tblGrid')
        gridcols = grid.findAll('gridCol')
        return float(gridcols[columnnumber].get('w'))/20.0

    def getColumnLeftOrRightBorder(self, columnnumber, side):
        borders = []
        row = self.getRow(0)
        while row is not None:
            cell = row.getCell(columnnumber)
            borders.append(cell.getBorder(side))
            row = row.getNextRow()

        result = 'nil'
        if len(borders)>0:
            result = borders[0]
            for b in borders:
                if b != result:
                    print("Table column", side, "borders not equal! Using first cell", side, "border.")

        return result

    def resolveColumnLeftOrRightBorder(self, columnnumber, side):

        # Get borders
        thisborder = self.getColumnLeftOrRightBorder(columnnumber, side)

        # Compare with previous row
        if side == 'left':
            otherside = 'right'
            othercolumn = columnnumber - 1
            if othercolumn < 0:
                othercolumn = None

        elif side == 'right':
            otherside = 'left'
            othercolumn = columnnumber + 1
            if othercolumn >= self.getNumberOfLogicalColumns():
                othercolumn = None
        else:
            print("Illegal border side!")

        if othercolumn is not None:
            otherborder = self.getColumnLeftOrRightBorder(othercolumn, otherside)
            if otherborder != thisborder:
                print("Column", side, "border not equal to other column", otherside, "border!")
                print("This column:", thisborder+"; other column:", otherborder)
                # Hierarchy: double over single over nil
                if otherborder == "double":
                    thisborder = "double"
                elif otherborder == "single":
                    thisborder = "single"
                print("Using", thisborder)

        return thisborder

    def hasText(self, disregardchars='', disregardpstyles=[]):
        tabletext = ''
        ps = self.findAll('p')
        for p in ps:
            if p.getStyle() not in disregardpstyles:
                ts = p.findAll('t')
                for t in ts:
                    tabletext += t.getText().strip()

        for c in tabletext:
            if c not in disregardchars:
                return True

        return False


class tblGrid(WordprocessingMLElement):
    """
    This element specifies the table grid for the current table. The table grid is a definition of the set of grid columns
    which define all of the shared vertical edges of the table, as well as default widths for each of these grid
    columns. These grid column widths are then used to determine the size of the table based on the table layout
    algorithm used.
    If the table grid is omitted, then a new grid shall be constructed from the actual contents of the table assuming
    that all grid columns have a width of 0.
    
    Parent elements: tbl
    
    Child elements: gridCol, tblGridChange
    """

    pass


class gridCol(WordprocessingMLElement):
    """
    This element specifies the presence and details about a single grid column within a table grid. A grid column is a
    logical column in a table used to specify the presence of a shared vertical edge in the table. When table cells are
    then added to this table, these shared edges (or grid columns, looking at the column between those shared
    edges) determine how table cells are placed into the table grid.
    
    Parent elements: tblGrid
    """

    pass


class tblPr(WordprocessingMLElement):
    """
    This element specifies the set of table-wide properties applied to the current table. These properties affect the
    appearance of all rows and cells within the parent table, but may be overridden by individual table-level
    exception, row, and cell level properties as defined by each property.
    
    Parent elements: tbl
    
    Child elements: bidiVisual, jc, shd, tblBorders, tblCellMar, tblCellSpacing, tblInd, tblLayout, tblLook, tblOverlap,
    tblpPr, tblPrChange, tblStyle, tblStyleColBandSize, tblStyleRowBandSize, tblW
    """

    pass


class tblW(WordprocessingMLElement):
    """
    This element specifies the preferred width for this table. This preferred width is used as part of the
    table layout algorithm specified by the tblLayout element - full description of the algorithm in the
    ST_TblLayout simple type.
    All widths in a table are considered preferred because:
    - The table must satisfy the shared columns as specified by the tblGrid element.
    - Two or more widths may have conflicting values for the width of the same grid column
    - The table layout algorithm may require a preference to be overridden
    This value is specified in the units applied via its type attribute. Any width value of type pct
    for this element shall be calculated relative to the text extents of the page (page width excluding margins).
    If this element is omitted, then the cell width shall be of type auto.
    """

    def getValue(self):
        w = self.get('w')
        t = self.get('type')
        if t == 'auto':
            return 'auto'
        elif t == 'dxa':
            return "{:.1f}".format(float(w)*0.017638889) # 1 twip = 0.017638889 mm
        elif t == 'nil':
            return 0
        elif t == 'pct':
            return "{:.1f}%".format(float(w)/50.0) # value in fiftieths of a percent
        return None


class shd(WordprocessingMLElement):
    """
    This element specifies the shading which shall be applied to the extents the current table. Similarly to paragraph
    shading, this shading shall be applied to the contents of the tab up to the table borders, regardless of the
    presence of text - unlike cell shading, table shading shall include any cell padding. This property shall be
    superseded by any cell-level shading via any table-level property exceptions; or on any cell in this row.
    This shading consists of three components:
    - Background Color (optional)
    - Pattern (optional)
    - Pattern Color
    The resulting shading is applied by setting the background color behind the paragraph, then applying the pattern
    color using the mask supplied by the pattern over that background.
    If this element is omitted, then the cells within this table shall have the shading specified by the associated
    table style. If no cell shading is specified in the style hierarchy, then the cells in this table shall not have
    any cell shading (i.e. they shall be transparent).
    """

    pass


class tr(WordprocessingMLElement):
    """
    This element specifies a single table row, which contains the table’s cells. Table rows in WordprocessingML are
    analogous to HTML tr elements.
    A tr element has one formatting child element, trPr (§2.4.78), which defines the properties for the row. Each
    unique property on the table row is specified by a child element of this element. As well, a table row can contain
    any valid row-level content, which allows for the use of table cells.
    If a table cell does not include at least one child element other than the row properties, then this document shall
    be considered corrupt.
    """

    def getNumberOfCells(self):
        return len(self.getCells())

    def getCells(self):
        return self.findAllInChildren('tc')

    ## AANPASSEN VANWEGE GRIDBEFORE?
    def getCellByLogicalColumnIndex(self, index):
        cs = self.getCells()
        curindex = 1 + self.getGridBefore()
        for c in cs:
            #curindex += 1
            if curindex == index:
                return c
            curindex += c.getColumnSpanning()
        return None

    def getCell(self, index):
        cs = self.getCells()
        if index >= len(cs):
            raise ValueError("Cell number exceeds row length!")
        return cs[index]

    def getRowProperties(self):
        return self.findFirst('trPr')

    def getGridBefore(self):
        trpr = self.getRowProperties()
        if trpr is not None:
            return trpr.getGridBefore()
        return 0

    def getPreviousRow(self):
        prev = self.getprevious()
        while prev is not None:
            if prev.tag == W+'tr':
                return prev
            prev = prev.getprevious()
        return None

    def getNextRow(self):
        next = self.getnext()
        while next is not None:
            if next.tag == W+'tr':
                return next
            next = next.getnext()
        return None

    def getTopOrBottomBorderStyle(self, side):
        tcs = self.getCells()
        borderstyles = []

        for c in tcs:
            borderstyles.append(c.resolveBorderStyle(side))

        result = 'nil'
        if len(borderstyles)>0:
            result = borderstyles[0]
            for b in borderstyles:
                if b != result:
                    print("Table row", side, "border styles not equal! Using first cell", side, "border style.")
        return result

    def getTopOrBottomBorderColor(self, side):
        tcs = self.getCells()
        bordercolors = []

        for c in tcs:
            bordercolors.append(c.resolveBorderColor(side))

        result = None
        if len(bordercolors)>0:
            result = bordercolors[0]
            for b in bordercolors:
                if b != result:
                    print("Table row", side, "border colors not equal! Using first cell", side, "border color.")
        return result

    def getRowColor(self):
        tcs = self.getCells()
        rowcolors = []

        for c in tcs:
            rowcolors.append(c.resolveCellColor())

        result = None
        if len(rowcolors)>0:
            result = rowcolors[0]
            for b in rowcolors:
                if b != result:
                    print("Table row colors not equal! Using first cell color.")
        return result


    def resolveTopOrBottomBorderStyle(self, side):
        # Get borders
        thisborderstyle = self.getTopOrBottomBorderStyle(side)

        # Compare with previous row
        if side == 'top':
            other = 'bottom'
            otherrow = self.getPreviousRow()
        elif side == 'bottom':
            other = 'top'
            otherrow = self.getNextRow()
        else:
            print("Illegal border side!")

        if otherrow is not None:
            otherborderstyle = otherrow.getTopOrBottomBorderStyle(other)
            if otherborderstyle != thisborderstyle:
                print("Row", side, "border style not equal to other row", other, "border style!")
                print("This row:", thisborderstyle+"; other row:", otherborderstyle)
                # Hierarchy: double over single over nil
                if otherborderstyle == "double":
                    thisborderstyle = "double"
                elif otherborderstyle == "single":
                    thisborderstyle = "single"
                print("Using", thisborderstyle)

        return thisborderstyle

    def resolveTopOrBottomBorderColor(self, side):
        # Get borders
        thisbordercolor = self.getTopOrBottomBorderColor(side)

        # Compare with previous row
        if side == 'top':
            other = 'bottom'
            otherrow = self.getPreviousRow()
        elif side == 'bottom':
            other = 'top'
            otherrow = self.getNextRow()
        else:
            print("Illegal border side!")

        if otherrow is not None:
            otherbordercolor = otherrow.getTopOrBottomBorderColor(other)
            if otherbordercolor != thisbordercolor:
                print("Row", side, "border color not equal to other row", other, "border color!")
                print("This row:", thisbordercolor, "; other row:", otherbordercolor)
                if thisbordercolor is None:
                    thisbordercolor = otherbordercolor
                print("Using", thisbordercolor)

        return thisbordercolor


class tc(WordprocessingMLElement):
    """
    This element specifies a single cell in a table row, which contains the table’s content. Table cells in
    WordprocessingML are analogous to HTML td elements.
    A tc element has one formatting child element, tcPr, which defines the properties for the cell. Each
    unique property on the table cell is specified by a child element of this element. As well, a table cell can contain
    any valid block-level content, which allows for the nesting of paragraphs and tables within table cells.
    If a table cell does not include at least one block-level element, then this document shall be considered corrupt.
    
    Parent element: tr
    
    Child elements: altChunk, bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml,
    """

    def getText(self):
        ps = self.findAll('p')
        ctext = ''
        for p in ps:
            for t in p.findAll('t'):
                ctext = ctext + t.getText()
        return ctext

    ## AANPASSEN VANWEGE GRIDBEFORE?
    def getLogicalColumn(self):
        r = self.getparent()
        tcs = r.getCells()
        lc = 1 + r.getGridBefore()
        for tc in tcs:
            if tc == self:
                break
            lc += tc.getColumnSpanning()
        return lc

    def getPreviousCell(self):
        prev = self.getprevious()
        while prev is not None:
            if prev.tag == W+'tc':
                return prev
            prev = prev.getprevious()
        return None

    def getNextCell(self):
        next = self.getnext()
        while next is not None:
            if next.tag == W+'tc':
                return next
            next = next.getnext()
        return None

    def getShadingDict(self):
        tcPr = self.findFirst('tcPr')
        if tcPr is not None:
            tcShading = tcPr.findFirst('shd')
            if tcShading is not None:
                val = tcShading.get('val')
                color = tcShading.get('color')
                fill = tcShading.get('fill')
                return {'val': val, 'color': color, 'fill': fill}
        return None

    def getBorderDict(self, side):
        tcPr = self.findFirst('tcPr')
        if tcPr is not None:
            tcBorders = tcPr.findFirst('tcBorders')
            if tcBorders is not None:
                b = tcBorders.findFirst(side)
                if b is not None:
                    val = b.get('val')
                    sz = b.get('sz')
                    space = b.get('space')
                    color = b.get('color')
                    return {'val': val, 'sz': sz, 'space': space, 'color': color}
        return None

    def resolveBorderDict(self, side):
        bdict = self.getBorderDict(side)
        if bdict is not None:
            return bdict
        else:
            # Query table properties
            return self.getTableBorderDict(side)

    def resolveShadingDict(self):
        sdict = self.getShadingDict()
        if sdict is not None:
            return sdict
        #cdict = self.getColumnShadingDict()
        #if cdict is not None:
        #	return cdict
        return self.getTableShadingDict()

    def getTableBorderDict(self, side):
        row = self.getparent()
        table = row.getparent()
        if (side == 'top') and (row is not table.getRow(0)):
            side = 'insideH'
        if (side == 'bottom') and (row is not table.getRow(-1)):
            side = 'insideH'
        if (side == 'left') and (self is not row.getCell(0)):
            side = 'insideV'
        if (side == 'right') and (self is not row.getCell(-1)):
            side = 'insideV'
        return table.getBorderDict(side)

    def getTableShadingDict(self):
        row = self.getparent()
        table = row.getparent()
        return table.getShadingDict()

    def resolveBorderStyle(self, side):
        bd = self.resolveBorderDict(side)
        if bd is not None:
            thisborderstyle = bd['val']
        else:
            thisborderstyle = 'nil'

        return thisborderstyle

    def resolveBorderColor(self, side):
        bd = self.resolveBorderDict(side)
        if bd is not None:
            thisbordercolor = bd['color']
        else:
            tbd = self.getTableBorderDict(side)
            if tbd is not None:
                thisbordercolor = tbd['color']
            else:
                thisbordercolor = None

        if thisbordercolor == 'auto':
            thisbordercolor = '000000'

        return thisbordercolor

    def getBackgroundColor(self):
        cellcolor = None

        sd = self.resolveShadingDict()
        if sd is not None:
            if sd['val'] == 'solid':
                cellcolor = sd['color']
            else:
                cellcolor = sd['fill']

        if cellcolor == 'auto':
            cellcolor = 'ffffff'

        return cellcolor

    def resolveCellTextColor(self):
        return None

    def getCellProperties(self):
        return self.findFirst('tcPr')

    def getColumnSpanning(self):
        cp = self.getCellProperties()
        if cp is not None:
            gs = cp.findFirst('gridSpan')
            if gs is not None:
                return int(gs.get('val'))
            hm = cp.findFirst('hMerge')
            if hm is not None:
                raise docxBase.DocxXMLError("Column spanning with hMerge!")
                #if hm.get('val') == 'restart':

        return 1

    def getNextVerticalCell(self):
        lcol = self.getLogicalColumn()
        thisrow = self.getparent()
        nextrow = thisrow.getNextRow()
        if nextrow is not None:
            return nextrow.getCellByLogicalColumnIndex(lcol)
        else:
            return None

    def getVMergeVal(self):
        cp = self.getCellProperties()
        if cp is not None:
            vm = cp.findFirst('vMerge')
            if vm is not None:
                vmv = vm.get('val')
                if vmv is not None:
                    return vmv
                else:
                    return 'continue'
        return None

    def getRowSpanning(self):
        vmv = self.getVMergeVal()
        if vmv == 'restart':
            return self.getNumberOfVMergedCells()
        return 1

    def getNumberOfVMergedCells(self):
        vmerge = 1
        c = self.getNextVerticalCell()
        while (c is not None) and (c.getVMergeVal() == 'continue'):
            c = c.getNextVerticalCell()
            vmerge += 1
        return vmerge


class tcPr(WordprocessingMLPropertiesElement):
    """
    This element specifies the set of properties which shall be applied a specific table cell. Each unique property is
    specified by a child element of this element. In any instance where there is a conflict between the table level,
    table-level exception, or row level properties with a corresponding table cell property, these properties shall
    overwrite the table or row wide properties.
    
    Parent element: tc
    
    Child elements: cellDel, cellIns, cellMerge, cnfStyle, gridSpan, hideMark, hMerge, noWrap, shd, tcBorders, tcFitText,
    tcMar, tcPrChange, tcW, textDirection, vAlign, vMerge
    """

    pass


class gridSpan(WordprocessingMLElement):
    """
    This element specifies the number of grid columns in the parent table's table grid which shall be spanned by the
    current cell. This property allows cells to have the appearance of being merged, as they span vertical boundaries
    of other cells in the table.
    If this element is omitted, then the number of grid units spanned by this cell shall be assumed to be one. If the
    number of grid units specified by the val attribute exceeds the size of the table grid, then the table grid shall
    augmented as needed to create the number of grid columns required.
    
    Parent element: tcPr
    """

    pass


class vMerge(WordprocessingMLElement):
    """
    This element specifies that this cell is part of a vertically merged set of cells in a table. The val attribute on this
    element determines how this cell is defined with respect to the previous cell in the table (i.e. does this cell
    continue the vertical merge or start a new merged group of cells).
    If this element is omitted, then this cell shall not be part of any vertically merged grouping of cells, and any
    vertically merged group of preceding cells shall be closed. If a vertically merged group of cells do not span the
    same set of grid columns, then this vertical merge is invalid.
    
    Parent element: tcPr
    """

    pass


class noWrap(WordprocessingMLElement):
    """
    This element specifies how this table cell shall be laid out when the parent table is displayed in a document. This
    setting only affects the behavior of the cell when the tblLayout for this row (§2.4.49; §2.4.50) is set to use the
    auto algorithm.
    This setting shall be interpreted in the context of the tcW element (§2.4.68) as follows:
    - If the table cell width has a type attribute value of fixed, then this element specifies that that this table
      cell shall never be smaller than that fixed value when other cells on the line are not at their absolute
      minimum width.
    - If the table cell width has a type attribute value of pct or auto, then this element specifies that when
      running the auto fit algorithm, the contents of that this table cell shall be treated as though they have
      no breaking characters (the contents should be treated as a single contiguous non-breaking string)
    If this element is omitted, then cell content shall be allowed to wrap (the cell may be shrunk as needed if it is a
    fixed preferred width value, and the contents shall be treated as having breaking characters if it is a percentage
    or automatic width value).
    
    Parent elements: tcPr
    """

    pass


class tcW(WordprocessingMLElement):
    """
    This element specifies the preferred width for this table cell. This preferred width is used as
    part of the table layout algorithm specified by the tblLayout element (§2.4.49; §2.4.50) - full
    description of the algorithm in the ST_TblLayout simple type (§2.18.94).
    All widths in a table are considered preferred because:
    - The table must satisfy the shared columns as specified by the tblGrid element (§2.4.44)
    - Two or more widths may have conflicting values for the width of the same grid column
    - The table layout algorithm (§2.18.94) may require a preference to be overridden
    This value is specified in the units applied via its type attribute. Any width value of type
    pct for this element shall be calculated relative to the overall width of the table.
    If this element is omitted, then the cell width shall be of type auto.
    
    Parent elements: tcPr
    """

    pass


class tblBorders(WordprocessingMLElement):
    """
    This element specifies the set of borders for the edges of the current table, using the six border types defined
    by its child elements.
    If the cell spacing for any row is non-zero as specified using the tblCellSpacing element (§2.4.41; §2.4.42; §2.4.43),
    then there is no border conflict and the table border (or table-level exception border, if one is specified) shall
    be displayed.
    If the cell spacing is zero, then there is a conflict [Example: Between the left border of all cells in the first
    column and the left border of the table. end example], which shall be resolved as follows:
    - If there is a cell border, then the cell border shall be displayed
    - If there is no cell border but there is a table-level exception border on this table row, then the table- level
      exception border shall be displayed
    - If there is no cell or table-level exception border, then the table border shall be displayed
    If this element is omitted, then this table shall have the borders specified by the associated table style. If
    no borders are specified in the style hierarchy, then this table shall not have any table borders.
    
    Parent element: tblPr
    
    Child elements: bottom, insideH, insideV, left, right, top
    """

    pass


class bottom(WordprocessingMLElement):
    """
    This element specifies the border which shall be displayed at the bottom of the current table. The appearance of this
    table border in the document shall be determined by the following settings:
    - The display of the border is subject to the conflict resolution algorithm defined by the tcBorders element (§2.4.63)
      and the tblBorders element (§2.4.37;§2.4.38)
    If this element is omitted, then the bottom of this table shall have the border specified by the associated table style.
    If no bottom border is specified in the style hierarchy, then this table shall not have a bottom border.
    
    Parent element: tblBorders
    """

    pass


class insideH(WordprocessingMLElement):
    pass


class insideV(WordprocessingMLElement):
    pass


class left(WordprocessingMLElement):
    pass


class right(WordprocessingMLElement):
    pass


class top(WordprocessingMLElement):
    pass


class tcBorders(WordprocessingMLElement):
    """
    Parent element: tcPr
    
    Child elements:
    """

    pass


class trPr(WordprocessingMLPropertiesElement):
    """
    This element specifies the set of row-level properties applied to the current table row. Each
    unique property is specified by a child element of this element. These properties affect the
    appearance of all cells in the current row within the parent table, but may be overridden by
    individual cell-level properties, as defined by each
    property.
    
    Parent element: tr
    
    Child elements: cantSplit, cnfStyle, del, divId, gridAfter, gridBefore, hidden, ins, jc,
    tblCellSpacing, tblHeader, trHeight, trPrChange, wAfter, wBefore
    """

    def getGridBefore(self):
        gb = self.findFirst('gridBefore')
        if gb is not None:
            return int(gb.get('val'))
        return 0


class trHeight(WordprocessingMLElement):
    """
    This element specifies the height of the current table row within the current table. This
    height shall be used to determine the resulting height of the table row, which may be
    absolute or relative (depending on its attribute values).
    If omitted, then the table row shall automatically resize its height to the height required
    by its contents (the equivalent of an hRule value of auto).
    
    Parent element: trPr
    """

    pass


class cantSplit(WordprocessingMLElement):
    """
    This element specifies whether the contents within the current cell shall be rendered on a
    single page. When displaying the contents of a table cell (such as the table cells in this
    specification), it is possible that a page break would fall within the contents of a table
    cell, causing the contents of that cell to be displayed across two different pages. If this
    property is set, then all contents of a table row shall be rendered on the same page by moving
    the start of the current row to the start of a new page if necessary. If the contents of this
    table row cannot fit on a single page, then this row shall start on a new page and flow onto
    multiple pages as necessary.
    If this element is not present, the default behavior is dictated by the setting in the associated
    table style. If this property is not specified in the style hierarchy, then this table row shall
    be allowed to split across multiple pages.
    
    Parent element: trPr
    """

    pass

### Misc classes


class sym(WordprocessingMLElement):
    """
    This element specifies the presence of a symbol character at the current location in the run’s
    content. A symbol character is a special character within a run’s content which does not use any of
    the run fonts specified in the rFonts element (or by the style hierarchy).
    Instead, this character shall be determined by pulling the character with the hexadecimal value
    specified in the char attribute from the font specified in the font attribute.
    
    Parent element: r
    """

    def getSymbolCode(self):
        c = self.get('char')
        return int(c, 16)

    def getSymbolFont(self):
        f = self.get('font')
        return f


class hyperlink(WordprocessingMLElement):
    """
    This element specifies the presence of a hyperlink at the current location in the document.
    
    Parent elements: customXml, fldSimple, hyperlink, p, sdtContent, smartTag
    
    Child elements: bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml, customXmlDelRangeEnd,
    customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart,
    customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, fldSimple, hyperlink, ins, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, permEnd, permStart, proofErr, r, sdt,
    smartTag, subDoc
    """

    pass


class bookmarkStart(WordprocessingMLElement):
    """
    This element specifies the start of a bookmark within a WordprocessingML document. This start marker is
    matched with the appropriately paired end marker by matching the value of the id attribute from the associated
    bookmarkEnd element.
    If no bookmarkEnd element exists subsequent to this element in document order with a matching id attribute
    value, then this element is ignored and no bookmark is present in the document with this name.
    If a bookmark begins and ends within a single table, it is possible for that bookmark to cover discontiguous parts
    of that table which are logically related (e.g. a single column in a table). This type of placement for a bookmark is
    accomplished (and described in detail) on the colFirst and colLast attributes on this element.
    
    Parent elements: body, comment, customXml, deg, del, den, docPartBody, e, endnote, fldSimple, fName, footnote, ftr, hdr,
    hyperlink, ins, lim, moveFrom, moveTo, num, oMath, p, rt, rubyBase, sdtContent, smartTag, sub, sup, tbl, tc, tr, txbxContent
    """

    def getBookmarkId(self):
        return self.get('id')

    def getBookmarkName(self):
        return self.get('name')

    def getBookmarkEnd(self):
        body = self.getroottree().find(W+'body')
        ends = body.findAll('bookmarkEnd')
        ebmnode = None

        for e in ends:
            if e.getBookmarkId() == self.getBookmarkId():
                ebmnode = e
                break

        if ebmnode is None:
            print("No bookmark end found, ignoring bookmark with id"+str(self.getBookmarkId()))

        return ebmnode

    def getBookmarkType(self, headingstyles, floatnames, captionstyles):
        self.headingP = None
        self.captionP = None
        self.footnoteID = None
        self.equationNode = None

        hasHeading = False
        hasFootnote = False
        hasEquation = False
        hasCaption = False

        ebmnode = self.getBookmarkEnd()
        if ebmnode is None:
            return None

        # Check ancestors
        for p in self.iterancestors(W+'p'):
            ps = p.getStyle()
            if not hasHeading:
                if ps in headingstyles:
                    hasHeading = True
                    self.headingP = p
            if not hasCaption:
                if ps in captionstyles:
                    hasCaption = True
                    self.captionP = p


        # Check elements between start and end
        n = self
        parent = self.getparent()
        while not (n == ebmnode):
            if not (n == self):
                if not hasFootnote:
                    if self.isFootnote(n):
                        hasFootnote = True
                if not hasHeading:
                    if self.isHeading(n, headingstyles):
                        hasHeading = True
                if not hasCaption:
                    if self.isCaption(n, floatnames, captionstyles):
                        hasCaption = True
                if not hasEquation:
                    if self.isEquation(n):
                        hasEquation = True

            if len(n)>0:
                parent = n
                n = n[0]
            elif n.getnext() is not None:
                n = n.getnext()
            else:
                while parent.getnext() is None:
                    parent = parent.getparent()
                n = parent.getnext()
                parent = n.getparent()

        if int(hasHeading)+int(hasFootnote)+int(hasCaption)+int(hasEquation) > 1:
            return ("Undefined", None)
        if hasHeading:
            return ('Heading', self.headingP)
        if hasCaption:
            return ('Caption', self.captionP)
        if hasFootnote:
            return ('Footnote', self.footnoteID)
        if hasEquation:
            return ('Equation', self.equationNode)
        return ('None', None)

    def isFootnote(self, node):
        if node.tag == W+'footnoteReference':
            self.footnoteID = node.getFootnoteId()
            return True
        return False

    def isCaption(self, node, floatnames, captionstyles):
        # also check caption paragraph style?
        #if node.tag == W+'fldSimple':
        #	fieldcodes = node.get('instr').split()
        #elif node.tag == W+'instrText':
        #	fieldcodes = node.text.split()
        #el
        if node.tag == W+'p':
            if node.getStyle() in captionstyles:
                self.captionP = node
                return True
        return False

        if fieldcodes[0] == 'SEQ' and fieldcodes[1] in floatnames:
            return True
        return False

    def isEquation(self, node):
        if node.tag == W+'fldSimple':
            fieldcodes = node.get('instr').split()
        elif node.tag == W+'fldChar':
            if node.isBegin():
                fieldcodes = node.getInstructionText().split()
            else:
                return False
        else:
            return False

        if fieldcodes[0] == 'SEQ' and fieldcodes[1] == 'Equation':
                self.equationNode = node
                return True
        return False

    def isHeading(self, node, headingstyles):
        if node.tag == W+'p':
            if node.getStyle() in headingstyles:
                self.headingP = node
                #print("Heading: "+str(node))
                return True
        return False


class bookmarkEnd(WordprocessingMLElement):
    """
    This element specifies the end of a bookmark within a WordprocessingML document. This end marker is matched with the
    appropriately paired start marker by matching the value of the id attribute from the associated bookmarkStart element.
    If no bookmarkStart element exists prior to this element in document order with a matching id attribute value, then
    this element is ignored and no bookmark is present in the document with this name.
    
    Parent elements: body, comment, customXml, deg, del, den, docPartBody, e, endnote, fldSimple, fName, footnote, ftr, hdr,
    hyperlink, ins, lim, moveFrom, moveTo, num, oMath, p, rt, rubyBase, sdtContent, smartTag, sub, sup, tbl, tc, tr, txbxContent
    """

    def getBookmarkId(self):
        return self.get('id')


class softHyphen(WordprocessingMLElement):
    """
    This element specifies that an optional hyphen character shall be placed at the current location in the run
    content. An optional hyphen is a character which may be used as a valid line breaking character for the current
    line of text when displaying this WordprocessingML content, using the following logic:
    When this character is not the character which is used to break the line, then it shall not change the
    normal display of text (it shall have zero width);
    When this character is the character used to break the line, it shall display using the hyphen-minus
    character within the display of text
    
    Parent element: r
    """

    pass


class br(WordprocessingMLElement):
    """
    This element specifies that a break shall be placed at the current location in the run content. A break is a special
    character which is used to override the normal line breaking that would be performed based on the normal layout of
    the document's contents.
    The behavior of this break character (the location where text shall be restarted after this break) shall be determined by
    its type and clear attribute values, described below.
    
    Parent element: r
    """

    def getType(self):
        return self.get("type")


class lastRenderedPageBreak(WordprocessingMLElement):
    """
    This element specifies that this position delimited the end of a page when this document was last saved by an application
    which paginates its content.
    
    Parent element: r
    """

    pass


class smartTag(WordprocessingMLElement):
    """
    This element specifies the presence of a smart tag around one or more inline structures (runs, images, fields, etc.)
    within a paragraph. The attributes on this element shall be used to specify the name and namespace URI of the
    current smart tag.
    
    Parent elemens: customXml, del, fldSimple, hyperlink, ins, moveFrom, moveTo, p, sdtContent, smartTag
    """

    def getElement(self):
        return self.get('element')


### Field codes


class fldChar(WordprocessingMLElement):
    """
    This element specifies the presence of a complex field character at the current location in the parent run. A
    complex field character is a special character which delimits the start and end of a complex field or separates its
    field codes from its current field result.
    A complex field is defined via the use of the two required complex field characters: a start character, which
    specifies the beginning of a complex field within the document content; and an end character, which specifies
    the end of a complex field. This syntax allows multiple fields to be embedded (or "nested") within each other in
    a document.
    As well, because a complex field may specify both its field codes and its current result within the document,
    these two items are separated by the optional separator character, which defines the end of the field codes and
    the beginning of the field contents. The omission of this character shall be used to specify that the contents of
    the field are entirely field codes (i.e. the field has no result).
    If a complex field character is located in an inappropriate location in a WordprocessingML document, then its
    presence shall be ignored and no field shall be present in the resulting document when displayed. Also, if a
    complex field is not closed before the end of a document story, then no field shall be generated and each individual
    run shall be processed as if the field characters did not exist (i.e. the contents of all field code run content
    shall not be displayed, and the field results shall be displayed as literal text).
    
    Parent element: r
    
    Child elements: ffData, fldData, numberingChange
    """

    def getType(self):
        return self.get("fldCharType")

    def isBegin(self):
        if self.getType() == 'begin':
            return True
        return False

    def isEnd(self):
        if self.getType() == 'end':
            return True
        return False

    def isSeparate(self):
        if self.getType() == 'separate':
            return True
        return False

    def getInstructionText(self):
        if not self.isBegin():
            return None
        itext = ""
        r = self.getparent()
        while r is not None:
            fc = r.findFirst('fldChar')
            if fc is not None and fc.isSeparate():
                break
            its = r.findAll('instrText')
            for i in its:
                itext += i.text
            r = r.getNextContiguousRun()
        return itext

    def getResultText(self):
        if not self.isBegin():
            return None
        inResult = False
        rtext = ""
        r = self.getparent()
        while r is not None:
            fc = r.findFirst('fldChar')
            if fc is not None and fc.isSeparate():
                inResult = True

            if fc is not None and fc.isEnd():
                break

            if inResult:
                rtext += r.getText()

            r = r.getNextContiguousRun()
        return rtext


class instrText(WordprocessingMLElement):
    """
    This element specifies that this run contains field codes within a complex field in the document.
    If this element is contained within a run which is not part of a complex field's field codes, then it and its contents
    should be treated as regular text. If this element is contained within a del element, then the document is invalid.
    
    Parent element: r
    """

    pass


class fldSimple(WordprocessingMLElement):
    """
    This element specifies the presence of a simple field at the current location in the document. The semantics of
    this field are defined via its field codes.
    
    Parent elements: customXml, fldSimple, hyperlink, p, sdtContent, smartTag
    
    Child elements: bookmarkEnd, bookmarkStart, commentRangeEnd, commentRangeStart, customXml, customXmlDelRangeEnd,
    customXmlDelRangeStart, customXmlInsRangeEnd, customXmlInsRangeStart, customXmlMoveFromRangeEnd, customXmlMoveFromRangeStart,
    customXmlMoveToRangeEnd, customXmlMoveToRangeStart, del, fldData, fldSimple, hyperlink, ins, moveFrom, moveFromRangeEnd,
    moveFromRangeStart, moveTo, moveToRangeEnd, moveToRangeStart, oMath, oMathPara, permEnd, permStart, proofErr, r, sdt,
    smartTag, subDoc
    """

    def getInstructionText(self):
        return self.get('instr')

    def getContainedText(self):
        ct = ""
        ts = self.findAll('t')
        for t in ts:
            ct += t.getText()
        return ct


class noProof(WordprocessingMLElement):
    """
    This element specifies that the contents of this run shall not report any errors when the document is scanned for
    spelling and grammar.
    If this element is not present, the default value is to leave the formatting applied at previous level in the style
    hierarchy. If this element is never applied in the style hierarchy, then spelling and grammar error shall not be
    suppressed on the contents of this run.
    
    Parent element: rPr
    """

    pass


class drawing(WordprocessingMLElement):
    """
    This element specifies that a DrawingML object is located at this position in the run’s contents. The layout
    properties of this DrawingML object are specified using the WordprocessingML Drawing syntax
    
    Parent elements: r
    
    Child elements: anchor, inline
    """

    def getGraphicsData(self):
        gd = self.findFirst('a:graphicData')
        if gd is None:
            raise ValueError("Drawing has no graphicsData")
        return gd

    def hasPic(self):
        gd = self.getGraphicsData()
        if gd.findFirst('pic:pic') is not None:
            return True
        return False

    def hasChart(self):
        gd = self.getGraphicsData()
        if gd.findFirst('c:chart') is not None:
            return True
        return False

    def getPic(self):
        if not self.hasPic():
            return None
        return self.getGraphicsData().findFirst('pic:pic')

    def getChart(self):
        if not self.hasChart():
            return None
        return self.getGraphicsData().findFirst('c:chart')

    def getImageReferenceId(self):
        if self.hasPic():
            blipnode = self.getPic().findFirst('a:blip')
            if blipnode is not None:
                return blipnode.get('r:embed')
        elif self.hasChart():
            chart = self.getGraphicsData().findFirst('c:chart')
            if chart is not None:
                return chart.get('r:id')
        return None

    def getImageSize(self):
        # Get size in cm
        extentnode = self.findFirst('wp:extent')

        imw = float(extentnode.get('cx', False).replace(',','.'))/360100
        imh = float(extentnode.get('cy', False).replace(',','.'))/360100
        return (imw, imh)


class pict(WordprocessingMLElement):
    """
    This element specifies that an object is located at this position in the run’s contents. The layout
    properties of this object are specified using the VML syntax
    
    Parent elements: r
    
    Child elements: Any element from the vml or office:office namespaces, control, movie
    """

    pass

### Numbering definitions


class numbering(WordprocessingMLElement):
    pass


class abstractNum(WordprocessingMLElement):
    pass


class lvl(WordprocessingMLElement):
    pass


class numFmt(WordprocessingMLElement):
    pass


class num(WordprocessingMLElement):
    pass


class abstractNumId(WordprocessingMLElement):
    pass

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = "{%s}" % W_NS
docxBase.NSMAP['w'] = W_NS

namespace = docxBase.lookup.get_namespace(W_NS)
namespace.update(vars())
namespace[None] = WordprocessingMLElement