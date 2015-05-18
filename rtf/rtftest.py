#!/usr/bin/env python
# RTF generator test program
#
# Demonstrates simple RTF document generation using the RTFDoc/TextDoc
# classes, which are based on the classes included with the POST project.

import RTFDoc
from TextDoc import *

ss = StyleSheet()
ps = ParagraphStyle()
ss.add_style("Default", ps)
pps = PaperStyle("Transcript", 27.94, 21.59)


doc = RTFDoc.RTFDoc(ss, pps, None)
doc.open("test.rtf")

doc.start_paragraph("Default")
doc.start_bold()
doc.write_text("Transcript Title")
doc.end_bold()
doc.end_paragraph()
#doc.font.set_underline(1)
doc.start_paragraph("Default")
doc.start_underline()
doc.write_text("Blah blah blah")
doc.end_underline()
doc.line_break()
doc.write_text("lkajdsf lkjsf lkjsdf lkjasf lkjf lkjsdf lkjsf lkjsfd lkjdsaf lkjsdaf lkjsaf dlkjsaf dlksadj flkjfd lkj lkjdsf alkdsaf jalkj lkjdas flfdlkfdlkjfsalkjfds lkjfdalkjfsdlk jflkjfdslkjaf ljk lkj lkjf lkjf lj fjkf kjfkjasdf ksjfd lkajdfslkd flkjadslfkj")
doc.end_paragraph()


doc.close()
