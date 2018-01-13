#    Copyright 2018 Roel Zinkstok
#
#    This file is part of zink.
#
#    zink is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    zink is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with zink.  If not, see <http://www.gnu.org/licenses/>.


from docxBase import DocX
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

