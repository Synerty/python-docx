#!/usr/bin/env python2.6
# -*- coding: utf-8 -*-
"""
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and
'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import logging
from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image
import zipfile
import re
import time
import os
    
log = logging.getLogger(__name__)

class Docx(object):
    ''' Open Docx Library
    
    This library has been converted to a class and altered to allow custom
    templates to be provided. Basically, you start with a template and then
    add the content.
    
    1) Provide a template, this is an existing word document.
    2) Call methods to add content to the document.
    3) Save the document.
    
    '''
    
    # The default template
    __templatePath = os.path.join(os.path.dirname(__file__), 'template.docx')
    
    # All Word prefixes / namespace matches used in document.xml & core.xml.
    # LXML doesn't actually use prefixes (just the real namespace) , but these
    # make it easier to copy Word output more easily.
    nsprefixes = {
        'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
        'o':  'urn:schemas-microsoft-com:office:office',
        've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        # Text Content
        'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        # Drawing
        'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
        'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'mv':  'urn:schemas-microsoft-com:mac:vml',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'v':   'urn:schemas-microsoft-com:vml',
        'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessing'
                'Drawing'),
        # Properties (core and extended)
        'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pr'
                'operties'),
        'dc':  'http://purl.org/dc/elements/1.1/',
        'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-'
                'properties'),
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        # Content Types
        'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
        # Package Relationships
        'r':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
               'ips'),
        'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
        # Dublin Core document properties
        'dcmitype': 'http://purl.org/dc/dcmitype/',
        'dcterms':  'http://purl.org/dc/terms/'}
    
    
    def __init__(self, template=None):
        self._relationshiplist = []
        self._document = None
        self._template = template if template else self.__templatePath
        self._media = {}
        self._coreprops = None
        self._appprops = None
        self._contentTypes = None
        self._webSettings = None
        
        if not os.path.isfile(self._template):
            raise Exception("template docx |%s|not found" % self._template)
        
        self._loaddocx()
        self._loadrels()
        self._loadmedia()
        self.coreproperties('none', 'none', 'none', '')
        self._initAppProps()
        self._initContentTypes()
        self._initWebSettings()
    
    def _loaddocx(self):
        '''Load the core document content into our xml "document" '''

        documentPath = 'word/document.xml'
        zf = zipfile.ZipFile(self._template)
        xmlcontent = zf.read(documentPath)
        self._document = etree.fromstring(xmlcontent)
        self._docbody = self._document.xpath('/w:document/w:body',
                                             namespaces=self.nsprefixes)[0]
    
    def _loadrels(self):
        '''Load the relationships content into our relationship list '''
        rl = []

        relsPath = 'word/_rels/document.xml.rels'
        zf = zipfile.ZipFile(self._template)
        
        if relsPath in zf.namelist():
          
            xmlcontent = zf.read(relsPath)
            rels = etree.fromstring(xmlcontent)
            
            for node in rels.getchildren():
                rl.append([node.get('Type'), node.get('Target')])
        
        else:
            # Fallback for when we're using the v0.2.1 version of the
            # default template
            rl = \
                [['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/numbering', 'numbering.xml'],
                 ['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/styles', 'styles.xml'],
                 ['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/settings', 'settings.xml'],
                 ['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/webSettings', 'webSettings.xml'],
                 ['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/fontTable', 'fontTable.xml'],
                 ['http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/theme', 'theme/theme1.xml']]
                
        self._relationshiplist = rl
    
    def _loadmedia(self):
        '''Load the relationships content into our relationship list '''
        zf = zipfile.ZipFile(self._template)
        media = {}
        prefix = 'word/media'
        for zipInfo in zf.infolist():
            name = zipInfo.filename
            if not name.startswith(prefix):
                continue
            
            name = name[len(prefix):]
            media[name] = zf.read(name)
            
        self._media = media
        
    def _initAppProps(self):
        """
        Create app-specific properties. See docproperties() for more common
        document properties.
    
        """
        appprops = etree.fromstring(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties x'
            'mlns="http://schemas.openxmlformats.org/officeDocument/2006/extended'
            '-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocum'
            'ent/2006/docPropsVTypes"></Properties>')
        props = \
            {'Template':             'Normal.dotm',
             'TotalTime':            '6',
             'Pages':                '1',
             'Words':                '83',
             'Characters':           '475',
             'Application':          'Microsoft Word 12.0.0',
             'DocSecurity':          '0',
             'Lines':                '12',
             'Paragraphs':           '8',
             'ScaleCrop':            'false',
             'LinksUpToDate':        'false',
             'CharactersWithSpaces': '583',
             'SharedDoc':            'false',
             'HyperlinksChanged':    'false',
             'AppVersion':           '12.0000'}
        for prop in props:
            appprops.append(self._makeelement(prop, tagtext=props[prop], nsprefix=None))

        self._appprops = appprops
        
    def _initContentTypes(self):
        types = etree.fromstring(
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/conten'
            't-types"></Types>')
        parts = {
            '/word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocu'
                                      'ment.theme+xml',
            '/word/fontTable.xml':    'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.fontTable+xml',
            '/docProps/core.xml':     'application/vnd.openxmlformats-package.co'
                                      're-properties+xml',
            '/docProps/app.xml':      'application/vnd.openxmlformats-officedocu'
                                      'ment.extended-properties+xml',
            '/word/document.xml':     'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.document.main+xml',
            '/word/settings.xml':     'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.settings+xml',
            '/word/numbering.xml':    'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.numbering+xml',
            '/word/styles.xml':       'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.styles+xml',
            '/word/webSettings.xml':  'application/vnd.openxmlformats-officedocu'
                                      'ment.wordprocessingml.webSettings+xml'}
        for part in parts:
            types.append(self._makeelement('Override', nsprefix=None,
                                     attributes={'PartName': part,
                                                 'ContentType': parts[part]}))
        # Add support for filetypes
        filetypes = {
            'gif':  'image/gif',
            'jpeg': 'image/jpeg',
            'jpg':  'image/jpeg',
            'png':  'image/png',
            'rels': 'application/vnd.openxmlformats-package.relationships+xml',
            'xml':  'application/xml'
        }
        for extension in filetypes:
            attrs = {
                'Extension':   extension,
                'ContentType': filetypes[extension]
            }
            default_elm = self._makeelement('Default', nsprefix=None, attributes=attrs)
            types.append(default_elm)
            
        self._contentTypes = types
        
    def _initWebSettings(self):
        '''Generate websettings'''
        web = self._makeelement('webSettings')
        web.append(self._makeelement('allowPNG'))
        web.append(self._makeelement('doNotSaveAsSingleFile'))
        self._webSettings = web
    
    def _makeelement(self, tagname, tagtext=None, nsprefix='w', attributes=None,
                    attrnsprefix=None):
        '''Create an element & return it'''
        # Deal with list of nsprefix by making namespacemap
        namespacemap = None
        if isinstance(nsprefix, list):
            namespacemap = {}
            for prefix in nsprefix:
                namespacemap[prefix] = self.nsprefixes[prefix]
            # FIXME: rest of code below expects a single prefix
            nsprefix = nsprefix[0]
        if nsprefix:
            namespace = '{' + self.nsprefixes[nsprefix] + '}'
        else:
            # For when namespace = None
            namespace = ''
        newelement = etree.Element(namespace + tagname, nsmap=namespacemap)
        # Add attributes with namespaces
        if attributes:
            # If they haven't bothered setting attribute namespace, use an empty
            # string (equivalent of no namespace)
            if not attrnsprefix:
                # Quick hack: it seems every element that has a 'w' nsprefix for
                # its tag uses the same prefix for it's attributes
                if nsprefix == 'w':
                    attributenamespace = namespace
                else:
                    attributenamespace = ''
            else:
                attributenamespace = '{' + self.nsprefixes[attrnsprefix] + '}'
    
            for tagattribute in attributes:
                newelement.set(attributenamespace + tagattribute,
                               attributes[tagattribute])
        if tagtext:
            newelement.text = tagtext
        return newelement
    
    
    def pagebreak(self, type='page', orient='portrait'):
        '''Insert a break, default 'page'.
        See http://openxmldeveloper.org/forums/thread/4075.aspx
        Return our page break element.'''
        # Need to enumerate different types of page breaks.
        validtypes = ['page', 'section']
        if type not in validtypes:
            tmpl = 'Page break style "%s" not implemented. Valid styles: %s.'
            raise ValueError(tmpl % (type, validtypes))
        pagebreak = self._makeelement('p')
        if type == 'page':
            run = self._makeelement('r')
            br = self._makeelement('br', attributes={'type': type})
            run.append(br)
            pagebreak.append(run)
        elif type == 'section':
            pPr = self._makeelement('pPr')
            sectPr = self._makeelement('sectPr')
            if orient == 'portrait':
                pgSz = self._makeelement('pgSz', attributes={'w': '12240', 'h': '15840'})
            elif orient == 'landscape':
                pgSz = self._makeelement('pgSz', attributes={'h': '12240', 'w': '15840',
                                                       'orient': 'landscape'})
            sectPr.append(pgSz)
            pPr.append(sectPr)
            pagebreak.append(pPr)
            
        self._docbody.append(pagebreak)
    
    
    def paragraph(self, paratext, style='BodyText', breakbefore=False, jc='left'):
        """
        Return a new paragraph element containing *paratext*. The paragraph's
        default style is 'Body Text', but a new style may be set using the
        *style* parameter.
    
        @param string jc: Paragraph alignment, possible values:
                          left, center, right, both (justified), ...
                          see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                          for a full list
    
        If *paratext* is a list, add a run for each (text, char_format_str)
        2-tuple in the list. char_format_str is a string containing one or more
        of the characters 'b', 'i', or 'u', meaning bold, italic, and underline
        respectively. For example:
    
            paratext = [
                ('some bold text', 'b'),
                ('some normal text', ''),
                ('some italic underlined text', 'iu')
            ]
        """
        # Make our elements
        paragraph = self._makeelement('p')
    
        if not isinstance(paratext, list):
            paratext = [(paratext, '')]
        text_tuples = []
        for pt in paratext:
            text, char_styles_str = (pt if isinstance(pt, (list, tuple))
                                     else (pt, ''))
            text_elm = self._makeelement('t', tagtext=text)
            if len(text.strip()) < len(text):
                text_elm.set('{http://www.w3.org/XML/1998/namespace}space',
                             'preserve')
            text_tuples.append([text_elm, char_styles_str])
        pPr = self._makeelement('pPr')
        pStyle = self._makeelement('pStyle', attributes={'val': style})
        pJc = self._makeelement('jc', attributes={'val': jc})
        pPr.append(pStyle)
        pPr.append(pJc)
    
        # Add the text to the run, and the run to the paragraph
        paragraph.append(pPr)
        for text_elm, char_styles_str in text_tuples:
            run = self._makeelement('r')
            rPr = self._makeelement('rPr')
            # Apply styles
            if 'b' in char_styles_str:
                b = self._makeelement('b')
                rPr.append(b)
            if 'i' in char_styles_str:
                i = self._makeelement('i')
                rPr.append(i)
            if 'u' in char_styles_str:
                u = self._makeelement('u', attributes={'val': 'single'})
                rPr.append(u)
            run.append(rPr)
            # Insert lastRenderedPageBreak for assistive technologies like
            # document narrators to know when a page break occurred.
            if breakbefore:
                lastRenderedPageBreak = self._makeelement('lastRenderedPageBreak')
                run.append(lastRenderedPageBreak)
            run.append(text_elm)
            paragraph.append(run)
        # Return the combined paragraph
        self._docbody.append(paragraph)
        return paragraph
    
    
    def contenttypes(self):
        return self._contentTypes
    
    
    def heading(self, headingtext, headinglevel, lang='en'):
        '''Make a new heading, return the heading element'''
        lmap = {'en': 'Heading', 'it': 'Titolo'}
        # Make our elements
        paragraph = self._makeelement('p')
        pr = self._makeelement('pPr')
        pStyle = self._makeelement(
            'pStyle', attributes={'val': lmap[lang] + str(headinglevel)})
        run = self._makeelement('r')
        text = self._makeelement('t', tagtext=headingtext)
        # Add the text the run, and the run to the paragraph
        pr.append(pStyle)
        run.append(text)
        paragraph.append(pr)
        paragraph.append(run)
        # Return the combined paragraph
        self._docbody.append(paragraph)
    
    
    def table(self, contents, heading=True, colw=None, cwunit='dxa', tblw=0,
              twunit='auto', borders={}, celstyle=None):
        """
        Return a table element based on specified parameters
    
        @param list contents: A list of lists describing contents. Every item in
                              the list can be a string or a valid XML element
                              itself. It can also be a list. In that case all the
                              listed elements will be merged into the cell.
        @param bool heading:  Tells whether first line should be treated as
                              heading or not
        @param list colw:     list of integer column widths specified in wunitS.
        @param str  cwunit:   Unit used for column width:
                                'pct'  : fiftieths of a percent
                                'dxa'  : twentieths of a point
                                'nil'  : no width
                                'auto' : automagically determined
        @param int  tblw:     Table width
        @param int  twunit:   Unit used for table width. Same possible values as
                              cwunit.
        @param dict borders:  Dictionary defining table border. Supported keys
                              are: 'top', 'left', 'bottom', 'right',
                              'insideH', 'insideV', 'all'.
                              When specified, the 'all' key has precedence over
                              others. Each key must define a dict of border
                              attributes:
                                color : The color of the border, in hex or
                                        'auto'
                                space : The space, measured in points
                                sz    : The size of the border, in eighths of
                                        a point
                                val   : The style of the border, see
                    http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
        @param list celstyle: Specify the style for each colum, list of dicts.
                              supported keys:
                              'align' : specify the alignment, see paragraph
                                        documentation.
        @return lxml.etree:   Generated XML etree element
        """
        table = self._makeelement('tbl')
        columns = len(contents[0])
        # Table properties
        tableprops = self._makeelement('tblPr')
        tablestyle = self._makeelement('tblStyle', attributes={'val': ''})
        tableprops.append(tablestyle)
        tablewidth = self._makeelement(
            'tblW', attributes={'w': str(tblw), 'type': str(twunit)})
        tableprops.append(tablewidth)
        if len(borders.keys()):
            tableborders = self._makeelement('tblBorders')
            for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                if b in borders.keys() or 'all' in borders.keys():
                    k = 'all' if 'all' in borders.keys() else b
                    attrs = {}
                    for a in borders[k].keys():
                        attrs[a] = unicode(borders[k][a])
                    borderelem = self._makeelement(b, attributes=attrs)
                    tableborders.append(borderelem)
            tableprops.append(tableborders)
        tablelook = self._makeelement('tblLook', attributes={'val': '0400'})
        tableprops.append(tablelook)
        table.append(tableprops)
        # Table Grid
        tablegrid = self._makeelement('tblGrid')
        for i in range(columns):
            attrs = {'w': str(colw[i]) if colw else '2390'}
            tablegrid.append(self._makeelement('gridCol', attributes=attrs))
        table.append(tablegrid)
        # Heading Row
        row = self._makeelement('tr')
        rowprops = self._makeelement('trPr')
        cnfStyle = self._makeelement('cnfStyle', attributes={'val': '000000100000'})
        rowprops.append(cnfStyle)
        row.append(rowprops)
        if heading:
            i = 0
            for heading in contents[0]:
                cell = self._makeelement('tc')
                # Cell properties
                cellprops = self._makeelement('tcPr')
                if colw:
                    wattr = {'w': str(colw[i]), 'type': cwunit}
                else:
                    wattr = {'w': '0', 'type': 'auto'}
                cellwidth = self._makeelement('tcW', attributes=wattr)
                cellstyle = self._makeelement('shd', attributes={'val': 'clear',
                                                           'color': 'auto',
                                                           'fill': 'FFFFFF',
                                                           'themeFill': 'text2',
                                                           'themeFillTint': '99'})
                cellprops.append(cellwidth)
                cellprops.append(cellstyle)
                cell.append(cellprops)
                # Paragraph (Content)
                if not isinstance(heading, (list, tuple)):
                    heading = [heading]
                for h in heading:
                    if isinstance(h, etree._Element):
                        cell.append(h)
                    else:
                        cell.append(self.paragraph(h, jc='center'))
                row.append(cell)
                i += 1
            table.append(row)
        # Contents Rows
        for contentrow in contents[1 if heading else 0:]:
            row = self._makeelement('tr')
            i = 0
            for content in contentrow:
                cell = self._makeelement('tc')
                # Properties
                cellprops = self._makeelement('tcPr')
                if colw:
                    wattr = {'w': str(colw[i]), 'type': cwunit}
                else:
                    wattr = {'w': '0', 'type': 'auto'}
                cellwidth = self._makeelement('tcW', attributes=wattr)
                cellprops.append(cellwidth)
                cell.append(cellprops)
                # Paragraph (Content)
                if not isinstance(content, (list, tuple)):
                    content = [content]
                for c in content:
                    if isinstance(c, etree._Element):
                        cell.append(c)
                    else:
                        if celstyle and 'align' in celstyle[i].keys():
                            align = celstyle[i]['align']
                        else:
                            align = 'left'
                        cell.append(self.paragraph(c, jc=align))
                row.append(cell)
                i += 1
            table.append(row)
        
        self._docbody.append(table)
        return table
    
    
    def picture(self, picfilepath,
            picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True,
            picname=None, overwrite=False):
        """
        Take a relationshiplist, picture file name, and return a paragraph
        containing the image and an updated relationshiplist.
        """
        # http://openxmldeveloper.org/articles/462.aspx
        # Create an image. Size may be specified, otherwise it will based on the
        # pixel size of image. Return a paragraph containing the picture'''
        # Copy the file into the media dir
        
        if not os.path.isfile(picfilepath):
          raise Exception('|%s| is not a valid file' % picfilepath)
        
        if picname == None:
          picname = os.path.basename(picfilepath)
          
        if not overwrite and picname in self._media:
          raise Exception('picname |%s| is already in this document' % picname)
          
        self._media[picname] = open(picfilepath, 'rb+').read()
        
        # Check if the user has specified a size
        if not pixelwidth or not pixelheight:
            # If not, get info from the picture itself
            pixelwidth, pixelheight = Image.open(picfilepath).size[0:2]
    
        # OpenXML measures on-screen objects in English Metric Units
        # 1cm = 36000 EMUs
        emuperpixel = 12700
        width = str(pixelwidth * emuperpixel)
        height = str(pixelheight * emuperpixel)
    
        # Set relationship ID to the first available
        picid = '2'
        picrelid = 'rId' + str(len(self._relationshiplist) + 1)
        self._relationshiplist.append([
            ('http://schemas.openxmlformats.org/officeDocument/2006/relationship'
             's/image'), 'media/' + picname])
    
        # There are 3 main elements inside a picture
        # 1. The Blipfill - specifies how the image fills the picture area
        #    (stretch, tile, etc.)
        blipfill = self._makeelement('blipFill', nsprefix='pic')
        blipfill.append(self._makeelement('blip', nsprefix='a', attrnsprefix='r',
                        attributes={'embed': picrelid}))
        stretch = self._makeelement('stretch', nsprefix='a')
        stretch.append(self._makeelement('fillRect', nsprefix='a'))
        blipfill.append(self._makeelement('srcRect', nsprefix='a'))
        blipfill.append(stretch)
    
        # 2. The non visual picture properties
        nvpicpr = self._makeelement('nvPicPr', nsprefix='pic')
        cnvpr = self._makeelement(
            'cNvPr', nsprefix='pic',
            attributes={'id': '0', 'name': 'Picture 1', 'descr': picname})
        nvpicpr.append(cnvpr)
        cnvpicpr = self._makeelement('cNvPicPr', nsprefix='pic')
        cnvpicpr.append(self._makeelement(
            'picLocks', nsprefix='a',
            attributes={'noChangeAspect': str(int(nochangeaspect)),
                        'noChangeArrowheads': str(int(nochangearrowheads))}))
        nvpicpr.append(cnvpicpr)
    
        # 3. The Shape properties
        sppr = self._makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
        xfrm = self._makeelement('xfrm', nsprefix='a')
        xfrm.append(self._makeelement(
            'off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
        xfrm.append(self._makeelement(
            'ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
        prstgeom = self._makeelement(
            'prstGeom', nsprefix='a', attributes={'prst': 'rect'})
        prstgeom.append(self._makeelement('avLst', nsprefix='a'))
        sppr.append(xfrm)
        sppr.append(prstgeom)
    
        # Add our 3 parts to the picture element
        pic = self._makeelement('pic', nsprefix='pic')
        pic.append(nvpicpr)
        pic.append(blipfill)
        pic.append(sppr)
    
        # Now make the supporting elements
        # The following sequence is just: make element, then add its children
        graphicdata = self._makeelement(
            'graphicData', nsprefix='a',
            attributes={'uri': ('http://schemas.openxmlformats.org/drawingml/200'
                                '6/picture')})
        graphicdata.append(pic)
        graphic = self._makeelement('graphic', nsprefix='a')
        graphic.append(graphicdata)
    
        framelocks = self._makeelement('graphicFrameLocks', nsprefix='a',
                                 attributes={'noChangeAspect': '1'})
        framepr = self._makeelement('cNvGraphicFramePr', nsprefix='wp')
        framepr.append(framelocks)
        docpr = self._makeelement('docPr', nsprefix='wp',
                            attributes={'id': picid, 'name': 'Picture 1',
                                        'descr': picdescription})
        effectextent = self._makeelement('effectExtent', nsprefix='wp',
                                   attributes={'l': '25400', 't': '0', 'r': '0',
                                               'b': '0'})
        extent = self._makeelement('extent', nsprefix='wp',
                             attributes={'cx': width, 'cy': height})
        inline = self._makeelement('inline', attributes={'distT': "0", 'distB': "0",
                                                   'distL': "0", 'distR': "0"},
                             nsprefix='wp')
        inline.append(extent)
        inline.append(effectextent)
        inline.append(docpr)
        inline.append(framepr)
        inline.append(graphic)
        drawing = self._makeelement('drawing')
        drawing.append(inline)
        run = self._makeelement('r')
        run.append(drawing)
        paragraph = self._makeelement('p')
        paragraph.append(run)
        
        self._docbody.append(paragraph)
    
    
    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        result = False
        searchre = re.compile(search)
        for element in self._document.iter():
            if element.tag == '{%s}t' % self.nsprefixes['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result
    
    
    def replace(self, search, replace):
        """
        Replace all occurences of string with a different string, return updated
        document
        """
        searchre = re.compile(search)
        for element in self._document.iter():
            if element.tag == '{%s}t' % self.nsprefixes['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        element.text = re.sub(search, replace, element.text)
    
    
    def _clean(self):
        """ Perform misc cleaning operations on documents.
            Returns cleaned document.
        """
    
        # Clean empty text and r tags
        for t in ('t', 'r'):
            rmlist = []
            for element in self._document.iter():
                if element.tag == '{%s}%s' % (self.nsprefixes['w'], t):
                    if not element.text and not len(element):
                        rmlist.append(element)
            for element in rmlist:
                element.getparent().remove(element)
    
    
    def _findTypeParent(self, element, tag):
        """ Finds fist parent of element of the given type
    
        @param object element: etree element
        @param string the tag parent to search for
    
        @return object element: the found parent or None when not found
        """
    
        p = element
        while True:
            p = p.getparent()
            if p.tag == tag:
                return p
    
        # Not found
        return None
    
    
    def AdvSearch(self, search, bs=3):
        '''Return set of all regex matches
    
        This is an advanced version of python-docx.search() that takes into
        account blocks of <bs> elements at a time.
    
        What it does:
        It searches the entire document body for text blocks.
        Since the text to search could be spawned across multiple text blocks,
        we need to adopt some sort of algorithm to handle this situation.
        The smaller matching group of blocks (up to bs) is then adopted.
        If the matching group has more than one block, blocks other than first
        are cleared and all the replacement text is put on first block.
    
        Examples:
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search : 'Hello,'
        output blocks : [ 'Hello,' ]
    
        original text blocks : [ 'Hel', 'lo', ' __', 'name', '__!' ]
        search : '(__[a-z]+__)'
        output blocks : [ '__name__' ]
    
        @param instance  document: The original document
        @param str       search: The text to search for (regexp)
                              append, or a list of etree elements
        @param int       bs: See above
    
        @return set      All occurences of search string
    
        '''
    
        # Compile the search regexp
        searchre = re.compile(search)
    
        matches = []
    
        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []
    
        for element in self._document.iter():
            if element.tag == '{%s}t' % self.nsprefixes['w']:  # t (text) elements
                if element.text:
                    # Add this element to searchels
                    searchels.append(element)
                    if len(searchels) > bs:
                        # Is searchels is too long, remove first elements
                        searchels.pop(0)
    
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # l = search lenght
                    # s = search start
                    # e = element IDs to merge
                    found = False
                    for l in range(1, len(searchels) + 1):
                        if found:
                            break
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s + l <= len(searchels):
                                e = range(s, s + l)
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text
    
                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    matches.append(match.group())
                                    found = True
        return set(matches)
    
    
    def advReplace(self, search, replace, bs=3):
        """
        Replace all occurences of string with a different string, return updated
        document
    
        This is a modified version of python-docx.replace() that takes into
        account blocks of <bs> elements at a time. The replace element can also
        be a string or an xml etree element.
    
        What it does:
        It searches the entire document body for text blocks.
        Then scan thos text blocks for replace.
        Since the text to search could be spawned across multiple text blocks,
        we need to adopt some sort of algorithm to handle this situation.
        The smaller matching group of blocks (up to bs) is then adopted.
        If the matching group has more than one block, blocks other than first
        are cleared and all the replacement text is put on first block.
    
        Examples:
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hello,' / 'Hi!'
        output blocks : [ 'Hi!', '', ' world!' ]
    
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hello, world' / 'Hi!'
        output blocks : [ 'Hi!!', '', '' ]
    
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hel' / 'Hal'
        output blocks : [ 'Hal', 'lo,', ' world!' ]
    
        @param instance  document: The original document
        @param str       search: The text to search for (regexp)
        @param mixed     replace: The replacement text or lxml.etree element to
                             append, or a list of etree elements
        @param int       bs: See above
    
        @return instance The document with replacement applied
    
        """
        # Enables debug output
        DEBUG = False
    
        # Compile the search regexp
        searchre = re.compile(search)
    
        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []
    
        for element in self._document.iter():
            if element.tag == '{%s}t' % self.nsprefixes['w']:  # t (text) elements
                if element.text:
                    # Add this element to searchels
                    searchels.append(element)
                    if len(searchels) > bs:
                        # Is searchels is too long, remove first elements
                        searchels.pop(0)
    
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # l = search lenght
                    # s = search start
                    # e = element IDs to merge
                    found = False
                    for l in range(1, len(searchels) + 1):
                        if found:
                            break
                        # print "slen:", l
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s + l <= len(searchels):
                                e = range(s, s + l)
                                # print "elems:", e
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text
    
                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    found = True
    
                                    # I've found something :)
                                    if DEBUG:
                                        log.debug("Found element!")
                                        log.debug("Search regexp: %s",
                                                  searchre.pattern)
                                        log.debug("Requested replacement: %s",
                                                  replace)
                                        log.debug("Matched text: %s", txtsearch)
                                        log.debug("Matched text (splitted): %s",
                                                  map(lambda i: i.text, searchels))
                                        log.debug("Matched at position: %s",
                                                  match.start())
                                        log.debug("matched in elements: %s", e)
                                        if isinstance(replace, etree._Element):
                                            log.debug("Will replace with XML CODE")
                                        elif isinstance(replace(list, tuple)):
                                            log.debug("Will replace with LIST OF"
                                                      " ELEMENTS")
                                        else:
                                            log.debug("Will replace with:",
                                                      re.sub(search, replace,
                                                             txtsearch))
    
                                    curlen = 0
                                    replaced = False
                                    for i in e:
                                        curlen += len(searchels[i].text)
                                        if curlen > match.start() and not replaced:
                                            # The match occurred in THIS element.
                                            # Puth in the whole replaced text
                                            if isinstance(replace, etree._Element):
                                                # Convert to a list and process
                                                # it later
                                                replace = [replace]
                                            if isinstance(replace, (list, tuple)):
                                                # I'm replacing with a list of
                                                # etree elements
                                                # clear the text in the tag and
                                                # append the element after the
                                                # parent paragraph
                                                # (because t elements cannot have
                                                # childs)
                                                p = self._findTypeParent(
                                                    searchels[i],
                                                    '{%s}p' % self.nsprefixes['w'])
                                                searchels[i].text = re.sub(
                                                    search, '', txtsearch)
                                                insindex = p.getparent().index(p) + 1
                                                for r in replace:
                                                    p.getparent().insert(
                                                        insindex, r)
                                                    insindex += 1
                                            else:
                                                # Replacing with pure text
                                                searchels[i].text = re.sub(
                                                    search, replace, txtsearch)
                                            replaced = True
                                            log.debug(
                                                "Replacing in element #: %s", i)
                                        else:
                                            # Clears the other text elements
                                            searchels[i].text = ''
    
    
    def getdocumenttext(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        paratextlist = []
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in self._document.iter():
            # Find p (paragraph) elements
            if element.tag == '{' + self.nsprefixes['w'] + '}p':
                paralist.append(element)
        # Since a single sentence might be spread over multiple text elements,
        # iterate through each paragraph, appending all text (t) children to that
        # paragraphs text.
        for para in paralist:
            paratext = u''
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{' + self.nsprefixes['w'] + '}t':
                    if element.text:
                        paratext = paratext + element.text
                elif element.tag == '{' + self.nsprefixes['w'] + '}tab':
                    paratext = paratext + '\t'
            # Add our completed paragraph text to the list of paragraph text
            if not len(paratext) == 0:
                paratextlist.append(paratext)
        return paratextlist
    
    
    def coreproperties(self, title, subject, creator, keywords, lastmodifiedby=None):
        """
        Create core properties (common document properties referred to in the
        'Dublin Core' specification). See appproperties() for other stuff.
        """
        coreprops = self._makeelement('coreProperties', nsprefix='cp')
        coreprops.append(self._makeelement('title', tagtext=title, nsprefix='dc'))
        coreprops.append(self._makeelement('subject', tagtext=subject, nsprefix='dc'))
        coreprops.append(self._makeelement('creator', tagtext=creator, nsprefix='dc'))
        coreprops.append(self._makeelement('keywords', tagtext=','.join(keywords),
                         nsprefix='cp'))
        if not lastmodifiedby:
            lastmodifiedby = creator
        coreprops.append(self._makeelement('lastModifiedBy', tagtext=lastmodifiedby,
                         nsprefix='cp'))
        coreprops.append(self._makeelement('revision', tagtext='1', nsprefix='cp'))
        coreprops.append(
            self._makeelement('category', tagtext='Examples', nsprefix='cp'))
        coreprops.append(
            self._makeelement('description', tagtext='Examples', nsprefix='dc'))
        currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
        # Document creation and modify times
        # Prob here: we have an attribute who name uses one namespace, and that
        # attribute's value uses another namespace.
        # We're creating the element from a string as a workaround...
        for doctime in ['created', 'modified']:
            elm_str = (
                '<dcterms:%s xmlns:xsi="http://www.w3.org/2001/XMLSchema-instanc'
                'e" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:'
                'W3CDTF">%s</dcterms:%s>'
            ) % (doctime, currenttime, doctime)
            coreprops.append(etree.fromstring(elm_str))
            
        self._coreprops = coreprops
        return self._coreprops
    
    
    def appproperties(self):

        return self._appprops
    
    def websettings(self):
        return self._webSettings
    
    
    def _genRelationshipsTree(self):
        '''Generate a Word relationships file'''
        # Default list of relationships
        # FIXME: using string hack instead of making element
        # relationships = self._makeelement('Relationships', nsprefix='pr')
        relationships = etree.fromstring(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006'
            '/relationships"></Relationships>')
        count = 0
        for relationship in self._relationshiplist:
            # Relationship IDs (rId) start at 1.
            rel_elm = self._makeelement('Relationship', nsprefix=None,
                                  attributes={'Id':     'rId' + str(count + 1),
                                              'Type':   relationship[0],
                                              'Target': relationship[1]}
                                  )
            relationships.append(rel_elm)
            count += 1
        return relationships
    
    
    def savedocx(self, output):
        '''Save a modified document'''
      
        self._clean()
        
        docxfile = zipfile.ZipFile(
            output, mode='w', compression=zipfile.ZIP_DEFLATED)
        
        templateFile = zipfile.ZipFile(self._template)
    
        # Serialize our trees into out zip file
        treesandfiles = {'word/document.xml' : self._document,
                         'docProps/core.xml' : self._coreprops,
                         'docProps/app.xml' : self._appprops,
                         '[Content_Types].xml' : self._contentTypes,
                         'word/webSettings.xml' : self._webSettings ,
                         'word/_rels/document.xml.rels' : self._genRelationshipsTree()}
        
        for path, tree in treesandfiles.items():
            log.info('Saving: %s' % path)
            treestring = etree.tostring(tree, pretty_print=True)
            docxfile.writestr(path, treestring)
    
        # Add & compress support files
        files_to_ignore = ['.DS_Store']  # nuisance from some os's
        for filename in templateFile.namelist():
            if (os.path.basename(filename) in files_to_ignore
                or filename in treesandfiles):
                continue
            log.info('Saving: %s', filename)
            docxfile.writestr(filename, templateFile.read(filename))
        log.info('Saved new file to: %r', output)
        docxfile.close()
        
