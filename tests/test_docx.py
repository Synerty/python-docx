#!/usr/bin/env python2.6
'''
Test docx module
'''
import lxml
from docx import *

def testopendocx():
    '''Ensure an etree element is returned'''
    if isinstance(opendocx('Welcome to the Python docx module.docx'),lxml.etree._Element):
        pass
    else:
        assert False
        
def testnewdocument():
    pass

def testmakeelement():
    '''Ensure custom elements get created'''
    testelement = makeelement('testname',tagattributes={'testattribute':'testvalue'},tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    '''Ensure paragraph creates p elements'''
    testpara = paragraph('paratext',style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
    
def testtable():
    '''Ensure tables make sense'''
    testtable = table([['A1','A2'],['B1','B2'],['C1','C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    #print testtable
    #assert testtable.xpath('/'+ns+':tr[2]/'+ns+':tc[2]/'+ns+':p/'+ns+':r/'+ns+':t').text == 'C2'
    print testtable.xpath('/'+ns+':tr[2]')
    pass