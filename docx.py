#!/usr/bin/env python2.6
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - 

MIT licensed - see 
'''

from lxml import etree
import zipfile
import re
import time

# Namespaces used for the test (document.xml)
docns = {
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }

# Namespaces used for document properties (core.xml)
propns={
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    }

def getns(nsdict,prefix):
    '''Given a dict to search, a namespace prefix to look for, return a formatted namespace'''
    return '{'+nsdict[prefix]+'}'

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)    
    return document

def newdocument():
    document = makeelement('document',tagattributes=docns)
    document.append(makeelement('body'))
    return document

def makeelement(tagname,tagtext=None,tagnamespace=getns(docns,'w'),tagattributes=None,attributenamespace=None):
    '''Create an element & return it'''  
    newelement = etree.Element(tagnamespace+tagname)
    # Add attributes with namespaces
    if tagattributes:
        # If they haven't bothered setting attribute namespace, use the same one as the tag
        if not attributenamespace:
            attributenamespace = tagnamespace    
        for tagattribute in tagattributes:
            newelement.set(attributenamespace+tagattribute, tagattributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement
    

def paragraph(paratext,style='BodyText'):
    '''Make a new paragraph element, containing a run, and some text. 
    Return the paragraph element.'''
    # Make our elements
    paragraph = makeelement('p')
    run = makeelement('r')
    text = makeelement('t',tagtext=paratext)
    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle',tagattributes={'val':style})
    pPr.append(pStyle)

                
    # Add the text the run, and the run to the paragraph
    run.append(text)    
    paragraph.append(pPr)    
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph


def heading(headingtext,headinglevel):
    '''Make a new heading, return the heading element'''
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle',tagattributes={'val':'Heading'+str(headinglevel)})    
    run = makeelement('r')
    text = makeelement('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph   


def table(contents):
    '''Get a list of lists, return a table'''
    table = makeelement('tbl')
    columns = len(contents[0][0])    
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle',tagattributes={'val':'ColorfulGrid-Accent1'})
    tablewidth = makeelement('tblW',tagattributes={'w':'0','type':'auto'})
    tablelook = makeelement('tblLook',tagattributes={'val':'0400'})
    for tableproperty in [tablestyle,tablewidth,tablelook]:
        tableprops.append(tableproperty)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = makeelement('tblGrid')
    for _ in range(columns):
        tablegrid.append(makeelement('gridCol',tagattributes={'gridCol':'2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle',tagattributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    for heading in contents[0]:
        cell = makeelement('tc')  
        # Cell properties  
        cellprops = makeelement('tcPr')
        cellwidth = makeelement('tcW',tagattributes={'w':'2390','type':'dxa'})
        cellstyle = makeelement('shd',tagattributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
        cellprops.append(cellwidth)
        cellprops.append(cellstyle)
        cell.append(cellprops)        
        # Paragraph (Content)
        cell.append(paragraph(heading))
        row.append(cell)
    table.append(row)            
    # Contents Rows   
    for contentrow in contents[1:]:
        row = makeelement('tr')     
        for content in contentrow:   
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            cellwidth = makeelement('tcW',tagattributes={'type':'dxa'})
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            cell.append(paragraph(content))
            row.append(cell)    
        table.append(row)   
    return table                 

def picture(filename):
    '''Create a pragraph containing an image - FIXME - not implemented yet'''
    # Word uses paragraphs to contain images
    # http://openxmldeveloper.org/articles/462.aspx
    resourceid = rId5
    newrelationship = makeelement('Relationship',tagattributes={'Id':resourceid,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'},Target=filename)
    newpara = makeelement('deleteme',style='BodyText')
    makeelement('drawing')
    makeelement('inline',tagattributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},tagnamespace=getns(docns,'wp'))
    makeelement('graphic',tagnamespace=getns(docns,'a'))
    makeelement('graphicData',tagnamespace=getns(docns,'a'))
    makeelement('pic',tagnamespace=getns(docns,'a'))
                          

def search(document,search):
    '''Search a document for a regex, return '''
    results = False
    searchre = re.compile(search)
    for element in document.iter():
        if element.tag == getns(docns,'w')+'t':
            if element.text:
                if searchre.match(element.text):
                    results = True
    return results

def replace(document,search,replace):
    '''Replace all occurences of string with a different string, return updated document'''
    newdocument = document
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == getns(docns,'w')+'t':
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument


def getdocumenttext(document):
    '''Get the contents of all text elements in the tree'''
    # Recursively get all elements beneath tree
    # Get each elements text attribute
    contents = ''
    for element in document.iter():
        if element.tag == getns(docns,'w')+'t':
            if element.text:
                contents = contents+element.text+'\n'
    return contents        

def docproperties(title,subject,creator,keywords,lastmodifiedby=None):
    '''Makes document properties. '''
    # OpenXML uses the term 'core' to refer to the 'Dublin Core' specification used to make the properties.  
    # FIXME: creating the coreProperties base element doesn't seem to be working - probably me needing to
    # know more about namespaces, so import from a string.
    docprops=etree.fromstring('''<cp:coreProperties 
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    xmlns:dcterms="http://purl.org/dc/terms/" 
    xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    </cp:coreProperties>''')    
    # BAD - trying to create the element normally
    # docprops = makeelement('coreProperties',tagnamespace=getns(propns,'cp'),tagattributes=propns,attributenamespace=getns(propns,'cp'))
    # generates this:
    '''<ns0:coreProperties 
    xmlns:ns0="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    ns0:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    ns0:dc="http://purl.org/dc/elements/1.1/" 
    ns0:dcterms="http://purl.org/dc/terms/" 
    ns0:dcmitype="http://purl.org/dc/dcmitype/"    
    ns0:xsi="http://www.w3.org/2001/XMLSchema-instance">'''
    docprops.append(makeelement('title',tagtext=title,tagnamespace=getns(propns,'dc')))
    docprops.append(makeelement('subject',tagtext=subject,tagnamespace=getns(propns,'dc')))
    docprops.append(makeelement('creator',tagtext=creator,tagnamespace=getns(propns,'dc')))
    docprops.append(makeelement('keywords',tagtext=','.join(keywords),tagnamespace=getns(propns,'cp')))    
    if not lastmodifiedby:
        lastmodifiedby = creator
    docprops.append(makeelement('lastModifiedBy',tagtext=lastmodifiedby,tagnamespace=getns(propns,'cp')))
    docprops.append(makeelement('revision',tagtext='1',tagnamespace=getns(propns,'cp')))
    docprops.append(makeelement('category',tagtext='Examples',tagnamespace=getns(propns,'cp')))
    docprops.append(makeelement('description',tagtext='Examples',tagnamespace=getns(propns,'dc')))
    currenttime = time.strftime('%Y-%m-%dT-%H:%M:%SZ')
    # FIXME - creating these items manually also fails - but we can live without them for now.
    #docprops.append(makeelement('created',tagattributes={'type':'dcterms:W3CDTF'},tagtext=currentime,tagnamespace=getns(propns,'dcterms'),attributenamespace=getns(propns,'xsi')))
    #docprops.append(makeelement('modified',tagattributes={'type':'dcterms:W3CDTF'},tagtext=currentime,tagnamespace=getns(propns,'dcterms'),attributenamespace=getns(propns,'xsi')))
    return docprops



def savedocx(document,properties,newfilename):
    '''Save a modified document'''
    newfile = zipfile.ZipFile(newfilename,mode='w')
    # Write our generated document
    documentstring = etree.tostring(document, pretty_print=True)
    newfile.writestr('word/document.xml',documentstring)
    # And it's properties
    propertiesstring = etree.tostring(properties, pretty_print=True)
    newfile.writestr('docProps/core.xml',propertiesstring)
    # Add support files
    for xmlfile in [ 
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/thumbnail.jpeg',
    'docProps/app.xml',
    'word/webSettings.xml',
    'word/_rels/document.xml.rels',
    'word/styles.xml',
    'word/numbering.xml',
    'word/theme/',
    'word/theme/theme1.xml',
    'word/settings.xml',
    'word/fontTable.xml']:
        newfile.write('template/'+xmlfile,xmlfile)
    print 'Saved new file to: '+newfilename
    return
    
