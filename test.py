from docx2xelatex import *
import os

fn = "/Users/rzinkstok/Dropbox/Zink/AA Current Work/20161114_TerSchure/conversie/Schure.docx"
print(os.path.exists(fn))

d = DocX(fn)

processor = docx2xelatex(d)
processor.process()
processor.writeLaTeXFiles()

