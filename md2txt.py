# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

# Copyright (C) 2018 Free Software Foundation
# Author : David Bateman <dbateman@free.fr>

# This script is a simple script to convert github markdown files to txt for
# inclusion in the distributed binaries

import os
import sys
import textwrap

if len(sys.argv) != 3:
    raise OSError("md2txt [mdFile] [txtFile]")

mdFile = sys.argv[1]
if not os.path.exists(mdFile) :
    raise OSError("Source markdown file does not exist")

with open(mdFile, "r") as mdFp:
    with open(sys.argv[2], "w") as txtFp:
        for line in mdFp:
            if line[0:4] == "####" :
                txtline = line[4:].lstrip()
                txtFp.write(os.linesep + txtline + "." * len(txtline) + os.linesep)
            elif line[0:3] == "###" :
                txtline = line[3:].lstrip()
                txtFp.write(os.linesep + txtline + "-" * len(txtline) + os.linesep)
            elif line[0:2] == "##" :
                txtline = line[2:].lstrip()
                txtFp.write(os.linesep + txtline + "=" * len(txtline) + os.linesep)
            elif line[0:1] == "#" :
                txtline = line[1:].lstrip()
                s = " " * int((80 - len(txtline)) / 2)
                txtFp.write(s + txtline + s + "=" * len(txtline) + os.linesep)
            elif line[0:3] == "---" :
                txtFp.write("\n")
            else :
                if "](" in line :
                    line = line.replace("[","").replace("]("," (")
                for txtline in textwrap.wrap(line, width=80):
                    txtFp.write(txtline + "\n")






