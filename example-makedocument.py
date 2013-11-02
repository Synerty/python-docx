#!/usr/bin/env python

"""
This file makes a .docx (Word 2007) file from scratch, showing off most of the
features of python-docx.

If you need to make documents from scratch, you can use this file as a basis
for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

from docx import Docx

if __name__ == '__main__':
    # Make a new document tree - this is the main part of a Word document
    docx = Docx()

    # Append two headings and a paragraph
    docx.heading("Welcome to Python's docx module", 1)
    docx.heading('Make and edit docx in 200 lines of pure Python', 2)
    docx.paragraph('The module was created when I was looking for a '
        'Python support for MS Word .doc files on PyPI and Stackoverflow. '
        'Unfortunately, the only solutions I could find used:')

    # Add a numbered list
    points = [ 'COM automation'
             , '.net or Java'
             , 'Automating OpenOffice or MS Office'
             ]
    for point in points:
        docx.paragraph(point, style='ListNumber')
    docx.paragraph('For those of us who prefer something simpler, I '
                          'made docx.')
    docx.heading('Making documents', 2)
    docx.paragraph('The docx module has the following features:')

    # Add some bullets
    points = ['Paragraphs', 'Bullets', 'Numbered lists',
              'Multiple levels of headings', 'Tables', 'Document Properties']
    for point in points:
        docx.paragraph(point, style='ListBullet')

    docx.paragraph('Tables are just lists of lists, like this:')
    # Append a table
    tbl_rows = [ ['A1', 'A2', 'A3']
               , ['B1', 'B2', 'B3']
               , ['C1', 'C2', 'C3']
               ]
    docx.table(tbl_rows)

    docx.heading('Editing documents', 2)
    docx.paragraph('Thanks to the awesomeness of the lxml module, '
                          'we can:')
    points = [ 'Search and replace'
             , 'Extract plain text of document'
             , 'Add and delete items anywhere within the document'
             ]
    for point in points:
        docx.paragraph(point, style='ListBullet')

    # Add an image
    docx.picture('image2.png', 'This is a test description')

    # Search and replace
    print 'Searching for something in a paragraph ...',
    if docx.search('the awesomeness'):
        print 'found it!'
    else:
        print 'nope.'

    print 'Searching for something in a heading ...',
    if docx.search('200 lines'):
        print 'found it!'
    else:
        print 'nope.'

    print 'Replacing ...',
    docx.replace('the awesomeness', 'the goshdarned awesomeness')
    print 'done.'

    # Add a pagebreak
    docx.pagebreak(type='page', orient='portrait')

    docx.heading('Ideas? Questions? Want to contribute?', 2)
    docx.paragraph('Email <python.docx@librelist.com>')

    docx.coreproperties(title='Python docx demo',
                        subject='A practical example of making docx from Python',
                        creator='Mike MacCana',
                        keywords= ['python', 'Office Open XML', 'Word'])

    # Save our document
    docx.savedocx('Welcome to the Python docx module.docx')

