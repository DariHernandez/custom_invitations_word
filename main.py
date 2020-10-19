#! python3
# ganerate word document with custom invitations

import os, docx

path = os.path.dirname(__file__)
txtPath = os.path.join (path, 'guest.txt') 
txtFile = open (txtPath, 'r')
docPath = os.path.join (path, 'invitations.docx') 
docFile = docx.Document()

# Counters for text 
linerCounter = 0
whiteSpaceCounter = 0

# Write each line
for line in txtFile.readlines(): 
    linerCounter += 1
    docFile.add_heading('I would be a pleasure to have the company of', 1)
    docFile.add_heading(line.strip(), 0)
    docFile.add_heading('at 11010 Memory Lane on the Evening of ', 2)
    docFile.add_heading('April 1st', 2)
    docFile.add_heading('at 7 o\'clock', 2)
    # COunt last paragraphs and ad white page
    docFile.paragraphs[4 * linerCounter + whiteSpaceCounter].runs[0].add_break(docx.enum.text.WD_BREAK.PAGE)
    whiteSpaceCounter += 1
    print ('Generating invitation for %s' % (line.strip()))

docFile.save(docPath)
