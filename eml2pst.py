# -*- coding: utf-8 -*-

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

# Copyright (C) 2016 Free Software Foundation
# Author : David Bateman <dbateman@free.fr>

# A very basic EML to PST conversion utility

import mapiex
import win32com.client
import os
import sys

# Get the directory to use for src and dest and derive a store name
if len(sys.argv) != 3:
    raise OSError("eml2pst [srcPath] [pstFile]")

srcPath = sys.argv[1]
if not os.path.exists(srcPath) :
    raise OSError("Source directory does not exist")
srcPath = srcPath.rstrip('/\\')

pst = sys.argv[2]
destPath = os.path.dirname(pst)
if not os.path.exists(destPath) :
    raise OSError("Destination directory does not exist")
storename = os.path.basename(pst)
if storename[-4:] != ".pst":
    raise OSError("PST file extension must be '.pst'")
storename = storename[:-4]

# Use COM/DDE to create the PST file
Outlook = win32com.client.Dispatch(r'Outlook.Application')

ns = Outlook.GetNamespace(r'MAPI')
print("Opening PST file - %s" % pst)
sys.stdout.flush()     
ns.AddStore(pst)
rootFolder = ns.Folders.GetLast()
rootFolder.Name = storename

# Open a MAPI instance for the importation of EML file
MAPI = mapiex.mapi()
MAPI.OpenMessageStore (storename)
rootfolder = MAPI.OpenRootFolder ()

# Walk through the directories in srcPath loading every EML file
c = 0
for dirpath, dirnames, files in os.walk(srcPath):
    if files == []:
        continue
    pstpath = dirpath[len(srcPath) + 1:]
    print("Importing EML files in %s" % pstpath)
    sys.stdout.flush()
    folder = rootfolder.CreateSubFolder (pstpath)
    for name in files:
        if name.lower().endswith('.eml'):
            c += 1
            if (c % 20) == 0:
                print("Importing message %d" % c)
                sys.stdout.flush()
            folder.ImportEML(os.path.join(dirpath, name))