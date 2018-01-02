from docx2xelatex import *
import os
import sys

if __name__ == "__main__":
    fn = sys.argv[1]
    if not os.path.exists(fn):
        print("File does not exist!")
        sys.exit()

    d = DocX(fn)
    processor = docx2xelatex(d)
    processor.process()
    processor.writeLaTeXFiles()

