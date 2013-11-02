#!/usr/bin/env python2.6
'''
Test docx module
'''
import os
import lxml
from docx import Docx

TEST_FILE = 'ShortTest.docx'
IMAGE1_FILE = 'image1.png'

# --- Setup & Support Functions ---
def setup_module():
    '''Set up test fixtures'''
    import shutil
    if IMAGE1_FILE not in os.listdir('.'):
        shutil.copyfile(os.path.join(os.path.pardir, IMAGE1_FILE), IMAGE1_FILE)
    testnewdocument()

def teardown_module():
    '''Tear down test fixtures'''
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def simpledoc():
    '''Make a docx (document, relationships) for use in other docx tests'''
    docx = Docx()
    docx.heading('Heading 1', 1)  
    docx.heading('Heading 2', 2)
    docx.paragraph('Paragraph 1')
    for point in ['List Item 1', 'List Item 2', 'List Item 3']:
        docx.paragraph(point, style='ListNumber')
    docx.pagebreak(type='page')
    docx.paragraph('Paragraph 2')
    docx.table([['A1', 'A2', 'A3'], ['B1', 'B2', 'B3'], ['C1', 'C2', 'C3']])
    docx.pagebreak(type='section', orient='portrait')
    docx.picture(IMAGE1_FILE, 'This is a test description')
    docx.pagebreak(type='section', orient='landscape')
    docx.paragraph('Paragraph 3')
    return docx


# --- Test Functions ---
def testsearchandreplace():
    '''Ensure search and replace functions work'''
    docx = simpledoc()
    assert docx.search('ing 1')
    assert docx.search('ing 2')
    assert docx.search('graph 3')
    assert docx.search('ist Item')
    assert docx.search('A1')
    if docx.search('Paragraph 2'): 
        docx.replace('Paragraph 2', 'Whacko 55') 
    assert docx.search('Whacko 55')
    
def testtextextraction():
    '''Ensure text can be pulled out of a document'''
    docx = Docx(TEST_FILE)
    paratextlist = docx.getdocumenttext()
    assert len(paratextlist) > 0

def testunsupportedpagebreak():
    '''Ensure unsupported page break types are trapped'''
    docx = Docx()
    try:
        docx.pagebreak(type='unsup')
    except ValueError:
        return  # passed
    assert False  # failed
    
def testnewdocument():
    '''Test that a new document can be created'''
    docx = Docx()
    docx.coreproperties('Python docx testnewdocument',
                        'A short example of making docx from Python',
                        'Alan Brooks',
                        ['python', 'Office Open XML', 'Word'])
    docx.savedocx(TEST_FILE)

def testopendocx():
    '''Ensure an etree element is returned'''
    docx = Docx(TEST_FILE)
    if isinstance(docx._document, lxml.etree._Element):
        pass
    else:
        assert False

def testmakeelement():
    '''Ensure custom elements get created'''
    docx = Docx()
    testelement = docx._makeelement('testname', attributes={'testattribute':'testvalue'}, tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    '''Ensure paragraph creates p elements'''
    docx = Docx()
    testpara = docx.paragraph('paratext', style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
    
def testtable():
    '''Ensure tables make sense'''
    docx = Docx()
    testtable = docx.table([['A1', 'A2'], ['B1', 'B2'], ['C1', 'C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    assert testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',
                           namespaces={'ns0':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text == 'B2'

if __name__ == '__main__':
    import nose
    nose.main()
