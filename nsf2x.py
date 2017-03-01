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

# Very loosely based on nlconverter (https://code.google.com/p/nlconverter/)
# by Hugues Bernard <hugues.bernard@gmail.com>

import gettext
import locale
import traceback
import tempfile
import datetime
import codecs
import os
import sys
import io
import ctypes
import win32com.client #NB : Calls to COM are starting with an uppercase
import pywintypes
import win32crypt
import win32cryptcon

try:
    # Python 3.x
    import tkinter
    import tkinter.ttk as ttk
except ImportError:
    # Python 2.7
    import Tkinter as tkinter
    import ttk

import mapiex

# This list should be extended to match regular install paths
notesDllPathList = [r'c:/notes', r'd:/notes', r'c:/program files/notes', r'd:/program files/notes',
                    r'c:/program files (x86)/notes', r'd:/program files (x86)/notes',
                    r'c:/program files/ibm/notes', r'd:/program files/ibm/notes',
                    r'c:/program files (x86)/ibm/notes', r'd:/program files (x86)/ibm/notes']

# Setup i8n
lang = locale.windows_locale.get(ctypes.windll.kernel32.GetUserDefaultLCID())
localedir = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0])), 'locale')
translate = gettext.translation('nsf2x', localedir, languages=[lang], fallback=True)
_ = translate.gettext
translate.install()

# The following classes are a means of creating a simple ENUM functionality
# Use list(range()) for Python 2.7 and 3.x compatibility
class Format:
    EML, MBOX, PST = list(range(3))

class EncryptionType:
    NONE, RC2CBC, DES, AES128, AES256 = list(range(5))

class SubdirectoryMBOX:
    NO, YES = list(range(2))

class ErrorLevel:
    NORMAL, ERROR, WARN, INFO = list(range(4))

class Exceptions:
    EX_1, EX_10, EX_100, EX_INF = list(range(4))

class NotesEntries(object):
    OPEN_RAW_RFC822_TEXT = ctypes.c_uint32(0x01000000)
    OPEN_RAW_MIME_PART = ctypes.c_uint32(0x02000000)
    OPEN_RAW_MIME	= ctypes.c_uint32(0x03000000) # OPEN_RAW_RFC822_TEXT | OPEN_RAW_MIME_PART
    nnotesdll = None
    hDb = ctypes.c_void_p(0)

    def __init__(self, fp=None):
        self.loaddll(fp)
        self.isLoaded(True, False)
        self.SetDLLReturnTypes()
        stat = self.nnotesdll.NotesInitExtended(0, ctypes.c_void_p(0))
        if stat != 0:
            raise NameError(_("NNOTES DLL can not be initialized (ErrorID %d)") % stat)

        # Throw an error if the DLL didn't load
        self.isLoaded(True, False)

    def __delete__(self, instance):
        self.NotesTerm()

    def loaddll(self, fp=None):
        if fp != None:
            if os.path.exists(fp):
                self.nnotesdll = ctypes.WinDLL(fp)
            else:
                self.nnotesdll = None
        else:
            self.nnotesdll = None
            try:
                # If we already have the COM/DDE interface to Notes, then nlsxbe.dll
                # is already loaded, so we can just try and get nnotes.dll leaving
                # Windows to search in its default search path
                self.nnotesdll = ctypes.WinDLL('nnotes.dll')
            except OSError:
                # Try harder
                for p in notesDllPathList:
                    fp = os.path.join(p, 'nnotes.dll')
                    if os.path.exists(fp):
                        self.nnotesdll = ctypes.WinDLL(fp)
                        break

    def isLoaded(self, raiseError=True, TestDb=True):
        if raiseError:
            if self.nnotesdll == None:
                raise NameError(_("NNOTES DLL not loaded"))
            elif TestDb and self.hDb == None:
                raise NameError(_("NNOTES DLL : Database not loaded"))
        else:
            return self.nnotesdll != None and self.hDb != None

    def SetDLLReturnTypes(self):
        self.nnotesdll.NotesInitExtended.restype = ctypes.c_uint16
        self.nnotesdll.NotesTerm.restype = ctypes.c_uint16
        self.nnotesdll.NSFDbOpen.restype = ctypes.c_uint16
        self.nnotesdll.NSFDbClose.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteOpenExt.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteOpenByUNID.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteClose.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteCopy.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteGetInfo.restype = None
        self.nnotesdll.NSFNoteIsSignedOrSealed.restype = ctypes.c_bool
        self.nnotesdll.NSFNoteDecrypt.restype = ctypes.c_uint16
        self.nnotesdll.NSFItemDelete.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteHasMIMEPart.restype = ctypes.c_bool
        self.nnotesdll.NSFNoteHasMIME.restype = ctypes.c_bool
        self.nnotesdll.NSFNoteHasComposite.restype = ctypes.c_bool
        self.nnotesdll.MMCreateConvControls.restype = ctypes.c_uint16
        self.nnotesdll.MMDestroyConvControls.restype = ctypes.c_uint16
        self.nnotesdll.MMSetMessageContentEncoding.restype = None
        self.nnotesdll.MIMEConvertCDParts.restype = ctypes.c_uint16
        self.nnotesdll.MIMEConvertMIMEPartCC.restype = ctypes.c_uint16
        self.nnotesdll.NSFNoteUpdate.restype = ctypes.c_uint16

    def NotesInitExtended(self, argc, argv):
        self.isLoaded(True, False)
        return self.nnotesdll.NotesInitExtended(argc, argv)

    def NotesTerm(self):
        self.isLoaded(True, False)
        return self.nnotesdll.NotesTerm()

    def NSFDbOpen(self, path):
        self.isLoaded(True, False)

        # Conversion UNICODE to LMBCS to allow Lotus to open databases with
        # accents in their names
        maxpath = 1024
        astr1 = path.encode('utf-8')
        astr2 = ctypes.create_string_buffer(maxpath)
        self.nnotesdll.OSTranslate(24, astr1, len(astr1), ctypes.byref(astr2), maxpath)

        return self.nnotesdll.NSFDbOpen(ctypes.c_char_p(astr2.value), ctypes.byref(self.hDb))

    def NSFDbClose(self):
        self.isLoaded()
        return self.nnotesdll.NSFDbClose(self.hDb)

    def NSFNoteCopy(self, hNote):
        self.isLoaded()
        hNoteNew = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFDbClose(hNote, ctypes.byref(hNoteNew))
        return retval, hNoteNew

    def NSFNoteOpenExt(self, nid, flags):
        self.isLoaded()
        hNote = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFNoteOpenExt(self.hDb, nid, flags, ctypes.byref(hNote))
        return retval, hNote

    def NSFNoteOpenByUNID(self, unid, flags):
        self.isLoaded()
        hNote = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFNoteOpenByUNID(self.hDb, unid, flags, ctypes.byref(hNote))
        return retval, hNote

    def NSFNoteClose(self, hNote):
        self.isLoaded()
        return self.nnotesdll.NSFNoteClose(hNote)

    def NSFNoteGetInfo(self, hNote, flags):
        self.isLoaded()
        retval = ctypes.c_uint16(0)
        self.nnotesdll.NSFNoteGetInfo(hNote, flags, ctypes.byref(retval))
        return retval

    def NSFNoteIsSignedOrSealed(self, hNote):
        self.isLoaded()
        isSigned = ctypes.c_bool(False)
        isSealed = ctypes.c_bool(False)
        retval = self.nnotesdll.NSFNoteIsSignedOrSealed(hNote, ctypes.byref(isSigned),
                                                        ctypes.byref(isSealed))
        return retval, isSigned.value, isSealed.value

    def NSFNoteDecrypt(self, hNote, flags):
        self.isLoaded()
        return self.nnotesdll.NSFNoteDecrypt(hNote, flags, ctypes.c_void_p(0))

    def NSFItemDelete(self, hNote, iname):
        self.isLoaded()
        return self.nnotesdll.NSFItemDelete(hNote, iname, len(iname))

    def NSFNoteHasMIMEPart(self, hNote):
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasMIMEPart(hNote)

    def NSFNoteHasMIME(self, hNote):
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasMIME(hNote)

    def NSFNoteHasComposite(self, hNote):
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasComposite(hNote)

    def MMCreateConvControls(self):
        self.isLoaded()
        hCC = ctypes.c_void_p(0)
        stat = self.nnotesdll.MMCreateConvControls(ctypes.byref(hCC))
        return(stat, hCC)

    def MMDestroyConvControls(self, hCC):
        self.isLoaded()
        return self.nnotesdll.MMDestroyConvControls(hCC)

    def MMSetMessageContentEncoding(self, hCC, flags):
        self.isLoaded()
        self.nnotesdll.MMSetMessageContentEncoding(hCC, flags)

    def MIMEConvertCDParts(self, hNote, bcanon, bisMime, hCC):
        self.isLoaded()
        return self.nnotesdll.MIMEConvertCDParts(hNote, bcanon, bisMime, hCC)

    def MIMEConvertMIMEPartsCC(self, hNote, bcanon, hCC):
        self.isLoaded()
        return self.nnotesdll.MIMEConvertCDParts(hNote, bcanon, hCC)

    def NSFNoteUpdate(self, hNote, flags):
        self.isLoaded()
        return self.nnotesdll.NSFNoteUpdate(hNote, flags)

class Gui(tkinter.Frame):
    """Basic Gui for NSF to EML, MBOX, PST export"""
    def __init__(self):

        # Setup the Tk frame including the manner in will the row/columns are
        # expanded. IE. Expand all columns equally, but only expand in height
        # the message area
        tkinter.Frame.__init__(self)
        self.master.title(_("Lotus Notes Converter"))
        self.master.grid_rowconfigure(4, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        self.master.grid_columnconfigure(3, weight=1)
        self.master.grid_columnconfigure(4, weight=1)

        self.nsfPath = "."
        self.destPath = os.path.join(os.path.expanduser('~'), 'Documents')
        self.checked = False
        self.Lotus = None
        self.running = False
        self.dialog = None
        self.certificate = None
        self.hCryptoProv = None

        # Initialize the default values of the Radio buttons
        self.Format = tkinter.IntVar()
        self.Format.set(Format.PST)
        self.Encrypt = tkinter.IntVar()
        self.Encrypt.set(EncryptionType.AES256)
        self.MBOXType = tkinter.IntVar()
        self.MBOXType.set(SubdirectoryMBOX.YES)
        self.ErrorLevel = tkinter.IntVar()
        self.ErrorLevel.set(ErrorLevel.ERROR)
        self.Exceptions = tkinter.IntVar()
        self.Exceptions.set(Exceptions.EX_100)

        # Lotus Password
        self.entryPassword = tkinter.Entry(self.master, relief=tkinter.GROOVE)
        self.entryPassword.insert(0, _("Enter Lotus Notes password"))
        self.entryPassword.grid(row=1, column=1, columnspan=2, sticky=tkinter.E+tkinter.W)
        self.entryPassword.bind("<FocusIn>", self.bindEntry)

        # Action button
        self.startButton = tkinter.Button(self.master, text=_("Open Sessions"),
                                          command=self.doConvert, relief=tkinter.GROOVE)
        self.startButton.grid(row=1, column=3, columnspan=2, sticky=tkinter.E+tkinter.W)

        # Conversion Type
        self.formatTypeEML = tkinter.Radiobutton(self.master, text="EML",
                                                 variable=self.Format, value=Format.EML)
        self.formatTypeEML.grid(row=2, column=1, sticky=tkinter.E+tkinter.W)
        self.formatTypeMBOX = tkinter.Radiobutton(self.master, text="MBOX",
                                                  variable=self.Format, value=Format.MBOX)
        self.formatTypeMBOX.grid(row=2, column=2, sticky=tkinter.E+tkinter.W)
        self.formatTypePST = tkinter.Radiobutton(self.master, text="PST",
                                                 variable=self.Format, value=Format.PST)
        self.formatTypePST.grid(row=2, column=3, sticky=tkinter.E+tkinter.W)

        # Options button
        self.optionsButton = tkinter.Button(self.master, text=_("Options"),
                                            command=self.doOptions, relief=tkinter.GROOVE,
                                            state=tkinter.DISABLED)
        self.optionsButton.grid(row=2, column=4, sticky=tkinter.E+tkinter.W)

        # Source chooser
        self.chooseNsfButton = tkinter.Button(self.master,
                                              text=_("Select Directory of SOURCE nsf files"),
                                              command=self.openSource, relief=tkinter.GROOVE,
                                              state=tkinter.DISABLED)
        self.chooseNsfButton.grid(row=3, column=1, columnspan=2, sticky=tkinter.E+tkinter.W)

        # Destination chooser
        self.chooseDestButton = tkinter.Button(self.master,
                                               text=_("Select Directory of DESTINATION files"),
                                               command=self.openDestination,
                                               relief=tkinter.GROOVE, state=tkinter.DISABLED)
        self.chooseDestButton.grid(row=3, column=3, columnspan=2, sticky=tkinter.E+tkinter.W)

        # Message Area
        frame = tkinter.Frame(self.master)
        frame.grid(row=4, column=1, columnspan=4, sticky=tkinter.E+tkinter.W+tkinter.N+tkinter.S)
        self.messageWidget = tkinter.Text(frame, width=80, height=20,
                                          state=tkinter.DISABLED, wrap=tkinter.NONE)
        scrollY = tkinter.Scrollbar(frame, orient=tkinter.VERTICAL,
                                    command=self.messageWidget.yview)
        self.messageWidget['yscrollcommand'] = scrollY.set
        scrollY.pack(side=tkinter.RIGHT, expand=tkinter.NO, fill=tkinter.Y)
        scrollX = tkinter.Scrollbar(frame, orient=tkinter.HORIZONTAL,
                                    command=self.messageWidget.xview)
        self.messageWidget['xscrollcommand'] = scrollX.set
        scrollX.pack(side=tkinter.BOTTOM, expand=tkinter.NO, fill=tkinter.X)

        self.messageWidget.pack(side=tkinter.RIGHT, expand=tkinter.YES, fill=tkinter.BOTH)
        self.log(ErrorLevel.NORMAL, _("Lotus Notes NSF file to EML, MBOX and PST file converter."))
        self.log(ErrorLevel.NORMAL, _("Contact dbateman@free.fr for more information.\n"))

    def openSource(self):
        dirname = self.tk.call('tk_chooseDirectory', '-initialdir', self.nsfPath,
                               '-mustexist', True)
        if dirname != "":
            self.nsfPath = dirname.replace('/', '\\')
            self.chooseNsfButton.config(text=_("Source directory is : %s") % self.nsfPath)

    def openDestination(self):
        dirname = self.tk.call('tk_chooseDirectory', '-initialdir',
                               self.destPath, '-mustexist', True)
        if dirname != "" and type(dirname) is not tuple and str(dirname) != "":
            self.destPath = dirname.replace('/', '\\')
            self.chooseDestButton.config(text=_("Destination directory is %s") % self.destPath)

    def bindEntry(self, dummy_event=None):
        """Blank the password field and set it in password mode"""
        self.entryPassword.delete(0, tkinter.END)
        self.entryPassword.config(show="*")
        self.entryPassword.unbind("<FocusIn>") #not needed anymore
        self.unchecked()

    def check(self):
        if self.Lotus != None:
            self.checked = True
            self.log(ErrorLevel.NORMAL, _("Connection to Notes established\n"))
        else:
            self.unchecked()
            self.log(ErrorLevel.ERROR, _("Check the Notes password\n"))
        return self.checked

    def unchecked(self):
        self.startButton.config(text=_("Open Sessions"))
        self.checked = False
        self.configPasswordEntry()

    def configStop(self, AllowButton=True, ActionText=_("Stop")):
        self.chooseNsfButton.config(state=tkinter.DISABLED)
        self.chooseDestButton.config(state=tkinter.DISABLED)
        self.entryPassword.config(state=tkinter.DISABLED)
        if AllowButton:
            self.startButton.config(text=ActionText, state=tkinter.NORMAL)
        else:
            self.startButton.config(text=ActionText, state=tkinter.DISABLED)
        self.optionsButton.config(state=tkinter.DISABLED)
        self.formatTypeEML.config(state=tkinter.DISABLED)
        self.formatTypeMBOX.config(state=tkinter.DISABLED)
        self.formatTypePST.config(state=tkinter.DISABLED)

    def configPasswordEntry(self):
        self.startButton.config(text=_("Open Sessions"), state=tkinter.NORMAL)
        self.chooseNsfButton.config(text=_("Select Directory of SOURCE nsf files"),
                                    state=tkinter.DISABLED)
        self.chooseDestButton.config(text=_("Select Directory of DESTINATION files"),
                                     state=tkinter.DISABLED)
        self.entryPassword.config(state=tkinter.NORMAL)
        self.formatTypeEML.config(state=tkinter.DISABLED)
        self.formatTypeMBOX.config(state=tkinter.DISABLED)
        self.formatTypePST.config(state=tkinter.DISABLED)
        self.optionsButton.config(state=tkinter.DISABLED)

    def configDirectoryEntry(self, SetDefaultPath=True):
        self.startButton.config(text=_("Convert"), state=tkinter.NORMAL)
        self.entryPassword.config(state=tkinter.DISABLED)
        self.formatTypeEML.config(state=tkinter.NORMAL)
        self.formatTypeMBOX.config(state=tkinter.NORMAL)
        self.formatTypePST.config(state=tkinter.NORMAL)
        self.optionsButton.config(state=tkinter.NORMAL)

        if SetDefaultPath:
            op = None
            try:
                op = os.path.join(os.path.dirname(self.Lotus.URLDatabase.FilePath), 'archive')
            except (pywintypes.com_error, OSError):
                try:
                    op = os.path.join(os.path.expanduser('~'), 'archive')
                except OSError:
                    op = None
            finally:
                if os.path.exists(op):
                    self.nsfPath = op
                else:
                    self.nsfPath = '.'

            sp = os.path.join(os.path.expanduser('~'), 'Documents')
            if os.path.exists(sp):
                self.destPath = sp
            else:
                self.destPath = '.'

        # TOBERM
        # This code is just to make my life while testing easier. Remove it eventually
        op = "C:\\Users\\C07056\\Documents\\temp"
        if os.path.exists(op):
            self.nsfPath = op
            self.destPath = op

        self.chooseNsfButton.config(text=_("Source directory is : %s") % self.nsfPath)
        self.chooseNsfButton.config(state=tkinter.NORMAL)
        self.chooseDestButton.config(text=_("Destination directory is %s") % self.destPath)
        self.chooseDestButton.config(state=tkinter.NORMAL)

    def doOptions(self):
        self.configStop(False, _("Convert"))

        self.dialog = tkinter.Toplevel(master=self.winfo_toplevel())
        self.dialog.title(_("NSF2X Options"))
        self.dialog.protocol("WM_DELETE_WINDOW", self.closeOptions)
        self.dialog.resizable(0, 0)

        L1 = tkinter.Label(self.dialog, text=_("Use different MBOXes for each sub-folder :"))
        L1.grid(row=1, column=1, columnspan=4, sticky=tkinter.W)

        R1 = tkinter.Radiobutton(self.dialog, text=_("No"), variable=self.MBOXType,
                                 value=SubdirectoryMBOX.NO)
        R1.grid(row=2, column=1, columnspan=2, sticky=tkinter.W)

        R2 = tkinter.Radiobutton(self.dialog, text=_("Yes"), variable=self.MBOXType,
                                 value=SubdirectoryMBOX.YES)
        R2.grid(row=2, column=3, columnspan=2, sticky=tkinter.W)

        ttk.Separator(self.dialog, orient=tkinter.HORIZONTAL).grid(row=3, columnspan=5,
                                                                   sticky=tkinter.E+tkinter.W)

        L2 = tkinter.Label(self.dialog, text=_("Re-encryption of encrypted Notes messages :"))
        L2.grid(row=4, column=1, columnspan=4, sticky=tkinter.W)

        R3 = tkinter.Radiobutton(self.dialog, text=_("None"), variable=self.Encrypt,
                                 value=EncryptionType.NONE)
        R3.grid(row=5, column=1, sticky=tkinter.W)

        R4 = tkinter.Radiobutton(self.dialog, text=_("RC2 40bit"), variable=self.Encrypt,
                                 value=EncryptionType.RC2CBC)
        R4.grid(row=5, column=2, sticky=tkinter.W)

        R5 = tkinter.Radiobutton(self.dialog, text=_("3DES 168bit"), variable=self.Encrypt,
                                 value=EncryptionType.DES)
        R5.grid(row=5, column=3, columnspan=2, sticky=tkinter.W)

        R6 = tkinter.Radiobutton(self.dialog, text=_("AES 128bit"), variable=self.Encrypt,
                                 value=EncryptionType.AES128)
        R6.grid(row=6, column=1, columnspan=2, sticky=tkinter.W)

        R7 = tkinter.Radiobutton(self.dialog, text=_("AES 256bit"), variable=self.Encrypt,
                                 value=EncryptionType.AES256)
        R7.grid(row=6, column=3, columnspan=2, sticky=tkinter.W)

        ttk.Separator(self.dialog, orient=tkinter.HORIZONTAL).grid(row=7, columnspan=5,
                                                                   sticky=tkinter.E+tkinter.W)

        L3 = tkinter.Label(self.dialog, text=_("Error logging level :"))
        L3.grid(row=8, column=1, columnspan=4, sticky=tkinter.W)

        R8 = tkinter.Radiobutton(self.dialog, text=_("Error"), variable=self.ErrorLevel,
                                 value=ErrorLevel.ERROR)
        R8.grid(row=9, column=1, sticky=tkinter.W)

        R9 = tkinter.Radiobutton(self.dialog, text=_("Warning"), variable=self.ErrorLevel,
                                 value=ErrorLevel.WARN)
        R9.grid(row=9, column=2, sticky=tkinter.W)

        R10 = tkinter.Radiobutton(self.dialog, text=_("Information"),
                                  variable=self.ErrorLevel, value=ErrorLevel.INFO)
        R10.grid(row=9, column=3, columnspan=2, sticky=tkinter.W)

        ttk.Separator(self.dialog, orient=tkinter.HORIZONTAL).grid(row=10, columnspan=5,
                                                                   sticky=tkinter.E+tkinter.W)

        L4 = tkinter.Label(self.dialog, text=_("Number of exceptions before giving up :"))
        L4.grid(row=11, column=1, columnspan=4, sticky=tkinter.W)

        R11 = tkinter.Radiobutton(self.dialog, text="1", variable=self.Exceptions,
                                  value=Exceptions.EX_1)
        R11.grid(row=12, column=1, sticky=tkinter.W)

        R12 = tkinter.Radiobutton(self.dialog, text="10", variable=self.Exceptions,
                                  value=Exceptions.EX_10)
        R12.grid(row=12, column=2, sticky=tkinter.W)

        R13 = tkinter.Radiobutton(self.dialog, text="100", variable=self.Exceptions,
                                  value=Exceptions.EX_100)
        R13.grid(row=12, column=3, sticky=tkinter.W)

        R14 = tkinter.Radiobutton(self.dialog, text=_("Infinite"),
                                  variable=self.Exceptions, value=Exceptions.EX_INF)
        R14.grid(row=12, column=4, sticky=tkinter.W)

        B1 = tkinter.Button(self.dialog, text=_("Close"), command=self.closeOptions,
                            relief=tkinter.GROOVE)
        B1.grid(row=13, column=2, columnspan=2, sticky=tkinter.E+tkinter.W)

        self.dialog.focus_force()

    def closeOptions(self):
        self.configDirectoryEntry(False)
        if self.dialog != None:
            self.dialog.destroy()

    def doConvert(self):
        if self.checked:
            if self.running:
                self.running = False
                self.configStop(False)
                self.log(ErrorLevel.NORMAL, _("Waiting for sub processes to terminate"))
            else:
                self.running = True
                self.configStop()
                self.master.after(0, self.doConvertDirectory())
        else: #Check if all is OK
            try:
                self.Lotus = win32com.client.Dispatch(r'Lotus.NotesSession')
                # Use rstrip to remove trailing whitespace as not part of the password
                self.Lotus.Initialize(self.entryPassword.get().rstrip())
                self.Lotus.ConvertMime = False
            except pywintypes.com_error as ex:
                self.log(ErrorLevel.ERROR, _("Error connecting to Lotus !"))
                self.log(ErrorLevel.ERROR, _("Exception %s :") % ex)
                # Try to force loading of Notes
                for p in notesDllPathList:
                    fp = os.path.join(p, 'nlsxbe.dll')
                    if os.path.exists(fp) and os.system('regsvr32 /s "%s"' % fp) == 0:
                        break
                self.Lotus = None

            self.check()
            if self.checked:
                self.configDirectoryEntry()

    def doConvertDirectory(self):
        tl = self.winfo_toplevel()
        self.log(ErrorLevel.NORMAL, _("Starting Convert : %s\n") % datetime.datetime.now())
        if self.Format.get() == Format.MBOX  and self.MBOXType.get() == SubdirectoryMBOX.NO:
            self.log(ErrorLevel.WARN, _("The MBOX file will not have the directory hierarchies present in NSF file\n"))

        for src in os.listdir(self.nsfPath):
            if not self.running:
                break

            abssrc = os.path.join(self.nsfPath, src)
            if os.path.isfile(abssrc) and src.lower().endswith('.nsf'):
                dest = src[:-4]
                try:
                    self.realConvert(src, dest)
                except (pywintypes.com_error, OSError) as ex:
                    self.log(ErrorLevel.ERROR, _("Error converting database %s") % src)
                    self.log(ErrorLevel.ERROR, _("Exception %s :") % ex)
                    self.log(ErrorLevel.ERROR, "%s" % traceback.format_exc())

        self.log(ErrorLevel.NORMAL, _("End of convert : %s\n") % datetime.datetime.now())
        tl.title(_("Lotus Notes Converter"))
        self.update()
        self.running = False
        self.configDirectoryEntry(False)

    def realConvert(self, src, dest):
        """Perform the translation from nsf to X"""
        c = 0 #document counter
        e = 0 #exception counter
        ac = 0 # all message count, though only an upper bounds as some documents not in folders
        tl = self.winfo_toplevel()

        # Setup the permitted number of exceptions
        if self.Exceptions.get() == Exceptions.EX_1:
            nex = 1
        elif self.Exceptions.get() == Exceptions.EX_10:
            nex = 10
        elif self.Exceptions.get() == Exceptions.EX_100:
            nex = 100
        else:
            nex = -1

        path = os.path.join(self.nsfPath, src)
        self.log(ErrorLevel.NORMAL, _("Converting : %s ") % path)

        if self.Lotus != None:
            try:
                dBNotes = self.Lotus.GetDatabase("", path)
                ac = dBNotes.AllDocuments.Count
            except pywintypes.com_error as ex:
                self.log(ErrorLevel.ERROR, _("Error connecting to Lotus !"))
                self.log(ErrorLevel.ERROR, _("Exception %s :") % ex)
        else:
            raise ValueError(_("Empty Lotus session"))

        if ac <= 0:
            raise ValueError(_("The database %s appears to be empty. Returning") % src)

        # Preconvert all messages to MIME before writing EML files as the
        # C DLL might not be finished saving the message before the COM
        # interface tries to access the MIME body. Also the call to mapiex.mapi()
        # must come after the conversion, as if it doesn't all the call to
        # MIMEConvertCDParts will raise a "File does not exist error (259)".
        # ?*#! -> Weird interaction MAPI to Notes
        # This also means that the NotesEntries class that loads nnotes.dll must
        # be called here rather that only once when starting NSF2X so that it is
        # reloaded after using mapîex.mapi() for multiple NSF files.
        #
        # If "File not found (259)" errors from MIMEConvertCDParts persist then
        # the call to "win32com.client.Dispatch(r'Lotus.NotesSession')" probably
        # needs to be in the method realConvert as well, though that will need
        # thought about reworking the UI. If after that there are still 259 errors
        # then NSF2X should be rewritten to force the user to relaunch after each
        # conversion, though that will prevent batch conversion of multiple NSF
        # files !!
        _NotesEntries = NotesEntries()
        stat = _NotesEntries.NSFDbOpen(path)
        if stat != 0:
            raise ValueError(_("ERROR : Can not open Lotus database %s with C API (ErrorID %d)") %
                             (path, stat))

        self.log(ErrorLevel.NORMAL, _("Starting MIME encoding of messages"))
        for fld in dBNotes.Views:
            if  not (fld.Name == "($Sent)" or fld.IsFolder) or fld.EntryCount <= 0:
                if fld.EntryCount > 0:
                    tl.title(_("Lotus Notes Converter - Phase 1/2 Converting MIME (%.1f%%)") %
                             float(10.*c/ac))
                    self.update()
                if not self.running:
                    return False
                continue
            doc = fld.GetFirstDocument()

            while doc and (nex < 0 or e < nex): #stop after XXX exceptions...
                if not self.running:
                    return False

                try:
                    if not self.ConvertToMIME(doc, _NotesEntries):
                        e += 1
                        self.log(ErrorLevel.ERROR, _("Can not convert message %d to MIME") % c)
                except (pywintypes.com_error, OSError) as ex:
                    e += 1
                    self.log(ErrorLevel.ERROR, _("Exception converting message %d to MIME : %s") %
                             (c, ex))

                doc = fld.GetNextDocument(doc)
                c += 1
                if (c % 20) == 0:
                    tl.title(_("Lotus Notes Converter - Phase 1/2 Converting MIME (%.1f%%)") %
                             float(10.*c/ac))
                    self.update()

        if e == nex:
            self.log(ErrorLevel.ERROR, _("Too many exceptions during MIME conversion. Stopping\n"))
            return False

        if c <= 0:
            raise ValueError(_("ERROR : The database %s appears to be empty. Returning") % src)

        f = None
        MAPIrootFolder = None

        if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.NO:
            mbox = os.path.join(self.destPath, (dest + ".mbox"))
            self.log(ErrorLevel.NORMAL, _("Opening MBOX file - %s") % mbox)
            f = open(mbox, "wb")
        elif self.Format.get() == Format.PST:
            pst = os.path.join(self.destPath, (dest + ".pst"))

            # Can't guarantee that MAPISVC.INF contains the service "MSPST MS" and so
            # can't use MAPI to create PST. This is now the only place the Outlook
            # Object Model is used, and it would be great to get rid of it.
            try:
                Outlook = win32com.client.Dispatch(r'Outlook.Application')
            except pywintypes.com_error as ex:
                self.log(ErrorLevel.ERROR, _("Could not connect to Outlook !"))
                self.log(ErrorLevel.ERROR, _("Exception %s :") % ex)
                Outlook = None
            ns = Outlook.GetNamespace(r'MAPI')
            self.log(ErrorLevel.NORMAL, _("Opening PST file - %s") % pst)
            ns.AddStore(pst)
            rootFolder = ns.Folders.GetLast()
            rootFolder.Name = dest

            # Reopen the message store created with OOM and only use MAPI from here
            # on out.
            try:
                MAPI = mapiex.mapi()
                MAPI.OpenMessageStore(dest)
                MAPIrootFolder = MAPI.OpenRootFolder()
            except Exception as ex:
                self.log(ErrorLevel.ERROR, _("Could not connect to MAPI !"))
                self.log(ErrorLevel.ERROR, _("Exception %s :") % ex)
                raise

        self.log(ErrorLevel.NORMAL, _("Starting importation of EML messages into mailbox"))
        ac = c # Update all message count
        c = 0
        e = 0
        for fld in dBNotes.Views:
            if  not (fld.Name == "($Sent)" or fld.IsFolder) or fld.EntryCount <= 0:
                if fld.EntryCount > 0:
                    tl.title(_("Lotus Notes Converter - Phase 2/2 Import Message %d of %d (%.1f%%)") % (c, ac, float(10.*(ac + 9.*c)/ac)))
                    self.update()
                if not self.running:
                    return False
                continue

            pstfld = None
            if self.Format.get() == Format.EML:
                if fld.Name == "($Sent)":
                    path = os.path.join(self.destPath, dest, _("Sent"))
                elif fld.Name == "($Inbox)":
                    path = os.path.join(self.destPath, dest, _("Inbox"))
                else:
                    path = os.path.join(self.destPath, dest, fld.Name)
                try:
                    if not os.path.exists(path):
                        os.makedirs(path, 0x755)
                        self.log(ErrorLevel.NORMAL, _("Creating directory %s") % path)
                except OSError as ex:
                    self.log(ErrorLevel.ERROR, _("Can not create directory %s") % path)
                    self.log(ErrorLevel.ERROR, "%s :" % ex)
                    continue
            elif self.Format.get() == Format.PST:
                if fld.Name == "($Sent)":
                    pstfld = MAPIrootFolder.CreateSubFolder(_("Sent"))
                elif fld.Name == "($Inbox)":
                    pstfld = MAPIrootFolder.CreateSubFolder(_("Inbox"))
                else:
                    pstfld = MAPIrootFolder.CreateSubFolder(fld.Name)

                if not pstfld:
                    self.log(ErrorLevel.ERROR, _("Could not open folder : %s") % fld.Name)
                    continue

            elif self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.YES:
                if fld.Name == "($Sent)":
                    mbox = os.path.join(self.destPath, dest, (_("Sent") + ".mbox"))
                elif fld.Name == "($Inbox)":
                    mbox = os.path.join(self.destPath, dest, (_("Inbox") + ".mbox"))
                else:
                    mbox = os.path.join(self.destPath, dest, (fld.Name + ".mbox"))

                try:
                    mboxdir = os.path.dirname(mbox)
                    if not os.path.exists(mboxdir):
                        os.makedirs(mboxdir, 0x755)
                        self.log(ErrorLevel.NORMAL, _("Creating directory %s") % mboxdir)
                except OSError as ex:
                    self.log(ErrorLevel.ERROR, _("Can not create directory %s") % mboxdir)
                    self.log(ErrorLevel.ERROR, "%s :" % ex)

                self.log(ErrorLevel.NORMAL, _("Opening MBOX file - %s") % mbox)
                f = open(mbox, "wb")

            doc = fld.GetFirstDocument()
            d = 1
            while doc and (nex < 0 or e < nex): #stop after XXX exceptions...
                if not self.running:
                    return False

                try:
                    eml = None

                    if doc.GetMIMEEntity("Body") == None:
                        subject = doc.GetFirstItem("Subject")
                        form = doc.GetFirstItem("Form")
                        if not form:
                            form = "None"
                        else:
                            form = form.Text
                        empty = False
                        if form in ("Appointment", "Task", "Notice", "Return Receipt",
                                    "Trace Report", "Delivery Report"):
                            # These are clearly not messages, so ok to ignore them
                            errlvl = ErrorLevel.WARN
                        else:
                            body = doc.GetFirstItem("Body")
                            if not body or body.ValueLength <= 0:
                                errlvl = ErrorLevel.WARN
                                empty = True
                            else:
                                errlvl = ErrorLevel.ERROR
                                e += 1

                        if empty:
                            self.log(errlvl, _("Ignoring message %d of form '%s' with empty body") % (c, form))
                        else:
                            self.log(errlvl, _("Ignoring message %d of form '%s' without MIME body") % (c, form))

                        if subject:
                            self.log(errlvl, _("#### Subject : %s") % subject.Text)

                        if errlvl == ErrorLevel.WARN:
                            self.log(errlvl, _("Skipping as probably not a message"))
                    else:
                        if self.Format.get() != Format.MBOX:
                            if self.Format.get() == Format.EML:
                                if fld.Name == "($Sent)":
                                    eml = os.path.join(self.destPath, dest, _("Sent"),
                                                       (str(d) + ".eml"))
                                elif fld.Name == "($Inbox)":
                                    eml = os.path.join(self.destPath, dest, _("Inbox"),
                                                       (str(d) + ".eml"))
                                else:
                                    eml = os.path.join(self.destPath, dest, fld.Name,
                                                       (str(d) + ".eml"))

                                # Need to treat as binary so that windows doesn't convert \n\r to \n\n\r
                                f = open(eml, "wb")
                            elif self.Format.get() == Format.PST:
                                (fd, eml) = tempfile.mkstemp(suffix=".eml")
                                f = os.fdopen(fd, "wb")

                        if self.WriteMIMEOutput(f, doc):
                            d += 1
                            if self.Format.get() == Format.PST:
                                f.close()
                                pstfld.ImportEML(eml)

                                # Done with the temporary EML file. Remove it
                                if eml != None:
                                    os.remove(eml)

                            elif self.Format.get() == Format.EML:
                                f.close()
                        else:
                            raise NameError(_("ERROR : Can not write Lotus MIME message to a file"))

                except (pywintypes.com_error, OSError) as ex:
                    e += 1 #count the exceptions
                    if self.Format.get() != Format.MBOX:
                        # File might already be closed and/or removed. So failure is ok
                        try:
                            f.close()
                        except OSError:
                            pass
                        try:
                            os.remove(eml)
                        except OSError:
                            pass
                    self.log(ErrorLevel.ERROR, _("Exception for message %d (%s) :") % (c, ex))
                    self.log(ErrorLevel.ERROR, "%s" % traceback.format_exc())
                    subject = doc.GetFirstItem("Subject")
                    if subject:
                        self.log(ErrorLevel.ERROR, _("#### Subject : %s") % subject.Text)

                finally:
                    c += 1
                    doc = fld.GetNextDocument(doc)

                    if self.Format.get() == Format.MBOX:
                        # MBOX is recognized by "\nFrom " string. So add a trailing \n to each message to ensure this format
                        f.write(b"\n")

                    if (c % 20) == 0:
                        tl.title(_("Lotus Notes Converter - Phase 2/2 Import Message %d of %d (%.1f%%)") % (c, ac, float(10.*(ac + 9.*c)/ac)))
                        self.update()

            if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.YES:
                f.close()

        # Alert user if there were too many exceptions
        if e == nex:
            self.log(ErrorLevel.ERROR, _("Too many exceptions during mail importation. Stopping"))

        if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.NO:
            f.close()
        self.log(ErrorLevel.NORMAL, _("Finished populating : %s") % dest)
        self.log(ErrorLevel.NORMAL, _("Exceptions: %d ... Documents OK : %d Untreated : %d\n") % (e, c - e, max(0, ac - c)))

        return True

    def ConvertToMIME(self, doc, _NotesEntries):
        # Check if NoteID is empty before continuing and give more informative
        # error message
        if doc.NoteID == None or doc.NoteID == '':
            self.log(ErrorLevel.ERROR, _("Notes message has empty NoteID"))
            return False

        # I'd really like to use doc.UniversalID here to open the file with
        # NSFNoteOpenByUNID. However, doc.UniversalID is a string and
        # NSFNoteOpenByUNID expects a struct and the conversion between the
        # two doesn't seem easy. Use doc.NoteID instead
        # stat, hNote = _NotesEntries.NSFNoteOpenByUNID(doc.UniversalID,
        #                                               _NotesEntries.OPEN_RAW_MIME)
        stat, hNote = _NotesEntries.NSFNoteOpenExt(ctypes.c_uint32(int(doc.NoteID, 16)),
                                                   _NotesEntries.OPEN_RAW_MIME)

        if stat != 0:
            self.log(ErrorLevel.ERROR, _("Can not open document id 0x%s (ErrorID : %d)") % (doc.NoteID, stat))
        else:
            try:
                # If present, $KeepPrivate will prevent conversion, so nuke the sucka
                tmp = doc.GetFirstItem("$KeepPrivate")
                if tmp != None:
                    self.log(ErrorLevel.INFO, _("Removing $KeepPrivate item from note id 0x%s") % doc.NoteID)
                    _NotesEntries.NSFItemDelete(hNote, "$KeepPrivate")

                # The C API identifies some unencrypted mail as "Sealed". These don't need
                # to be unencrypted to allow conversion to MIME.
                enc = doc.GetFirstItem("Encrypt")
                if enc != None and enc.Text == '1':
                    # if the note is encrypted, try to decrypt it. If that fails
                    #(e.g., we don't have the key), then we can't convert to MIME
                    # (we don't care about the signature)
                    stat, isSigned, isSealed = _NotesEntries.NSFNoteIsSignedOrSealed(hNote)
                    if isSealed:
                        self.log(ErrorLevel.INFO, _("Document note id 0x%s is encrypted.") % doc.NoteID)
                        DECRYPT_ATTACHMENTS_IN_PLACE = ctypes.c_uint16(1)
                        stat = _NotesEntries.NSFNoteDecrypt(hNote, DECRYPT_ATTACHMENTS_IN_PLACE)
                    if isSigned:
                        self.log(ErrorLevel.INFO, _("Document note id 0x%s is signed") % doc.NoteID)

                        if stat != 0:
                            self.log(ErrorLevel.ERROR, _("Document note id 0x%s is encrypted, cannot be converted.") % doc.NoteID)

                if stat == 0:
                    # if the note is already in mime format, we don't have to convert
                    if not _NotesEntries.NSFNoteHasMIMEPart(hNote):
                        stat, hCC = _NotesEntries.MMCreateConvControls()
                        if stat == 0:
                            _NotesEntries.MMSetMessageContentEncoding(hCC, 2) # html w/images & attachments

                            # NOTE_FLAG_CANONICAL = 0x4000 see nsfnote.h
                            _NOTE_FLAGS = ctypes.c_uint16(7)
                            bCanonical = (_NotesEntries.NSFNoteGetInfo(hNote, _NOTE_FLAGS).value) & 0x4000 != 0
                            bIsMime = _NotesEntries.NSFNoteHasMIMEPart(hNote)
                            stat = _NotesEntries.MIMEConvertCDParts(hNote, bCanonical, bIsMime, hCC)

                            if stat == 14941:
                                self.log(ErrorLevel.INFO, _("MIMEConvertCDParts : Error converting note id 0x%s to MIME type text/html") % doc.NoteID)
                                self.log(ErrorLevel.INFO, _("MIMEConvertCDParts : Attempting to convert to text/plain"))
                                _NotesEntries.MMSetMessageContentEncoding(hCC, 1)
                                stat = _NotesEntries.MIMEConvertCDParts(hNote, bCanonical,
                                                                        bIsMime, hCC)

                            if stat == 0:
                                UPDATE_FORCE = ctypes.c_uint16(1)
                                stat = _NotesEntries.NSFNoteUpdate(hNote, UPDATE_FORCE)
                                if stat != 0:
                                    self.log(ErrorLevel.ERROR,
                                             _("Error calling NSFNoteUpdate(%d)") % stat)
                            else:
                                self.log(ErrorLevel.ERROR,
                                         _("Error calling MIMEConvertCDParts(%d)") % stat)

                            _NotesEntries.MMDestroyConvControls(hCC)
                        else:
                            self.log(ErrorLevel.ERROR,
                                     _("Error calling MMCreateConvControls(%d)") % stat)

                if hNote != None:
                    _NotesEntries.NSFNoteClose(hNote)
            except:
                if hNote != None:
                    # Ensure Note is closed and then re-raise the exception
                    _NotesEntries.NSFNoteClose(hNote)
                raise

        return stat == 0

    def WriteMIMEHeader(self, f, mime):
        if mime != None:
            headers = mime.Headers
            encoding = mime.Encoding

            # if it's a binary part, force it to b64
            if encoding == 1730 or encoding == 1729:
                # MIMEEntity.ENC_IDENTITY_BINARY and MIMEEntity.ENC_IDENTITY_8BIT
                mime.EncodeContent(1727)  # MIMEEntity.ENC_BASE64
                headers = mime.Headers

            # Place the From and Date fields first to simplify conversion to MBOX format
            if self.Format.get() == Format.MBOX:
                content = mime.GetSomeHeaders(['From'], True)
                if content.startswith('From: '):
                    _from = content[6:]
                elif content.startswith('From:'):
                    _from = content[5:]
                else:
                    _from = content
                if _from.endswith('\n'):
                    _from = _from[:-1]
                content = mime.GetSomeHeaders(['Date'], True)
                if content.startswith('Date: '):
                    _date = content[6:]
                elif content.startswith('Date:'):
                    _date = content[5:]
                else:
                    _date = content
                if _date.endswith('\n'):
                    _date = _date[:-1]
                mboxheader = 'From ' + _from + ' ' + _date+ '\n'
                f.write(mboxheader.encode('utf-8'))

            # message envelope. If no MIME-Version header, add one
            if "MIME-Version:" not in headers:
                f.write(b"MIME-Version: 1.0\n")

            # Write the rest of the headers, but exclude the MIME content-type to be placed last
            content = mime.GetSomeHeaders(["Content-type"], False)
            # Some of the text might be in utf-8 so give it special treatment
            f.write(content.encode('utf-8'))
            if not content.endswith('\n'):
                f.write(b'\n')

    def WriteMIMEChildren(self, f, mime, first):
        """Iterate over children of MIME object"""
        if mime != None:
            contentType = mime.ContentType
            headers = mime.Headers
            encoding = mime.Encoding

            # if it's a binary part, force it to b64
            if encoding == 1730 or encoding == 1729:
                # MIMEEntity.ENC_IDENTITY_BINARY and MIMEEntity.ENC_IDENTITY_8BIT
                mime.DecodeContent()
                mime.EncodeContent(1727)  # MIMEEntity.ENC_BASE64
                headers = mime.Headers

            if first:
                content = mime.GetSomeHeaders(["Content-type"], True)
                f.write(content.encode('utf-8'))
                if not content.endswith('\n'):
                    f.write(b'\n')
            else:
                f.write(headers.encode('utf-8'))
                if not headers.endswith('\n'):
                    f.write(b'\n')

            f.write(b'\n')
            content = mime.ContentAsText
            if content != None:
                f.write(content.encode('utf-8'))
                if not content.endswith('\n'):
                    f.write(b'\n')

            f.flush()

            if contentType.startswith("multipart"):
                try:
                    # The preamble attribute might not exist
                    content = mime.preamble
                    if content != "":
                        f.write(content.encode('utf-8'))
                        if not content.endswith('\n'):
                            f.write(b'\n')
                except AttributeError:
                    pass

                child = mime.GetFirstChildEntity()
                while child != None:
                    content = child.BoundaryStart
                    f.write(content.encode('utf-8'))
                    if not content.endswith('\n'):
                        f.write(b'\n')

                    self.WriteMIMEChildren(f, child, False)

                    content = child.BoundaryEnd
                    f.write(content.encode('utf-8'))
                    if not content.endswith('\n'):
                        f.write(b'\n')

                    child = child.GetNextSibling()

    def WriteMIMEOutput(self, f_mime, doc):
        """Write MIME Output to EML file"""
        if doc != None:
            # Get first Body item with a MIME encoding
            mime = doc.GetMIMEEntity("Body")
            if mime != None:
                self.WriteMIMEHeader(f_mime, mime)
                if self.Encrypt.get() == EncryptionType.NONE:
                    self.WriteMIMEChildren(f_mime, mime, True)
                else:
                    enc = doc.GetFirstItem("Encrypt")
                    if enc != None and enc.Text == '1':
                        # See https://msdn.microsoft.com/en-us/library/windows/desktop/aa382376(v=vs.85).aspx
                        # Note that the PROV_RSA_AES provider supplies RC2, RC4 and
                        # AES encryption whereas as the PROV_RSA_FULL provider only
                        # gives RC2 and RC4 encryption. Try all possible combinations
                        # of providers to try and get a valid provider. Don't try and
                        # create a new provider however as we want a key that the user
                        # actually uses.
                        if not self.hCryptoProv:
                            # Loop through the various provider names, that are
                            # associated with PROV_RSA_AES
                            for prov in (win32cryptcon.MS_ENH_RSA_AES_PROV, None):
                                try:
                                    self.hCryptoProv = win32crypt.CryptAcquireContext(None, prov, win32cryptcon.PROV_RSA_AES, win32cryptcon.CRYPT_SILENT)
                                    break
                                except OSError as ex:
                                    self.log(ErrorLevel.ERROR, _("Exception : %s"), ex)

                            if not self.hCryptoProv:
                                if self.Encrypt.get() == EncryptionType.AES128 or self.Encrypt.get() == EncryptionType.AES256:
                                    self.log(ErrorLevel.ERROR, _("Windows cryptographic provider does not support AES encryption"))
                                    self.log(ErrorLevel.ERROR, _("Falling back to 3DES 168bit encryption"))
                                    self.Encrypt.set(EncryptionType.DES)

                                # Loop through the various provider names, that
                                # are associated with PROV_RSA_FULL
                                for prov in (win32cryptcon.MS_ENHANCED_PROV,
                                             win32cryptcon.MS_STRONG_PROV,
                                             win32cryptcon.MS_DEF_PROV, None):
                                    try:
                                        self.hCryptoProv = win32crypt.CryptAcquireContext(None, prov, win32cryptcon.PROV_RSA_FULL, win32cryptcon.CRYPT_SILENT)
                                        break
                                    except OSError as ex:
                                        self.log(ErrorLevel.ERROR, _("Exception : %s"), ex)

                            if not self.hCryptoProv:
                                self.log(ErrorLevel.ERROR, 
                                         _("Can not open Windows cryptographic provider"))

                        if self.hCryptoProv and not self.certificate:
                            hstorehandle = win32crypt.CertOpenSystemStore("MY", self.hCryptoProv)

                            for cert in hstorehandle.CertEnumCertificatesInStore():
                                try:
                                    (certtype, dummy_privcert) = cert.CryptAcquireCertificatePrivateKey(win32cryptcon.CRYPT_ACQUIRE_SILENT_FLAG)
                                    if certtype == win32cryptcon.AT_KEYEXCHANGE:
                                        # Ok we have the users key as we can access both
                                        # the public and private keys and the key is flagged
                                        # for use with Exchange
                                        self.certificate = cert
                                        break
                                except OSError:
                                    pass

                            if not self.certificate:
                                self.log(ErrorLevel.ERROR,
                                         _("Could not obtain the users Exchange certificate."))

                        if not self.hCryptoProv or not self.certificate:
                            self.log(ErrorLevel.ERROR, _("Disabling all encryption !!"))
                            self.WriteMIMEChildren(f_mime, mime, True)
                            self.Encrypt.set(EncryptionType.NONE)
                        else:
                            f_smime = io.BytesIO()
                            self.WriteMIMEChildren(f_smime, mime, True)

                            encodingtype = win32cryptcon.PKCS_7_ASN_ENCODING | win32cryptcon.X509_ASN_ENCODING

                            if self.Encrypt.get() == EncryptionType.RC2CBC:
                                encryptalgorithm = {"ObjId" :
                                                    win32cryptcon.szOID_RSA_RC2CBC,
                                                    "Parameters" : None}
                            elif self.Encrypt.get() == EncryptionType.DES:
                                encryptalgorithm = {"ObjId" :
                                                    win32cryptcon.szOID_RSA_DES_EDE3_CBC,
                                                    "Parameters" : None}
                            elif self.Encrypt.get() == EncryptionType.AES128:
                                # Why does win32cryptcon not define szOID_NIST_AES128_CBC
                                # and szOID_NIST_AES256_CBC ???
                                # szOID_NIST_AES128_CBC = "2.16.840.1.101.3.4.1.2"
                                # szOID_NIST_AES256_CBC = "2.16.840.1.101.3.4.1.42"
                                encryptalgorithm = {"ObjId" : "2.16.840.1.101.3.4.1.2",
                                                    "Parameters" : None}
                            elif self.Encrypt.get() == EncryptionType.AES256:
                                encryptalgorithm = {"ObjId" : "2.16.840.1.101.3.4.1.42",
                                                    "Parameters" : None}
                            else:
                                # This shouldn't be possible
                                raise NameError(_("ERROR : Unrecognised encryption selected"))

                            encryptparams = {"MsgEncodingType" : encodingtype, "CryptProv" :
                                             self.hCryptoProv, "ContentEncryptionAlgorithm" :
                                             encryptalgorithm}
                            blob = win32crypt.CryptEncryptMessage(encryptparams,
                                                                  [self.certificate],
                                                                  f_smime.getvalue())

                            f_mime.write(b'Content-Type: application/x-pkcs7-mime;smime-type=enveloped-data;name="smime.p7m"\n')
                            f_mime.write(b'Content-Transfer-Encoding: base64\n')
                            f_mime.write(b'Content-Disposition: attachment;filename="smime.p7m"\n')
                            f_mime.write(b'\n')

                            f_mime.write(codecs.encode(blob, "base64"))
                            f_smime.close()
                    else:
                        self.WriteMIMEChildren(f_mime, mime, True)
                return True
            else:
                self.log(ErrorLevel.WARN, _("Message 0x%s has no MIME body") % doc.NoteID)
                self.log(ErrorLevel.WARN, _("Type : %d") % doc.GetFirstItem("Body").Type)
                self.log(ErrorLevel.WARN, _("Subject : %s") % doc.GetFirstItem("Subject").Text)
        return False

    def log(self, errlvl, message="", newline=True):
        """Error logging function"""
        if errlvl == ErrorLevel.NORMAL:
            if self.ErrorLevel.get() >= ErrorLevel.NORMAL:
                message = _("INFO : ") + message
            else:
                return
        elif errlvl == ErrorLevel.ERROR:
            if self.ErrorLevel.get() >= ErrorLevel.ERROR:
                message = _("ERROR : ") + message
            else:
                return
        elif errlvl == ErrorLevel.WARN:
            if self.ErrorLevel.get() >= ErrorLevel.WARN:
                message = _("WARN : ") + message
            else:
                return
        elif errlvl == ErrorLevel.INFO:
            if self.ErrorLevel.get() >= ErrorLevel.INFO:
                message = _("INFO : ") + message
            else:
                return
        else:
            message = _("ERROR : Unrecognised Error Level given to log function")

        self.messageWidget.config(state=tkinter.NORMAL)
        if newline:
            self.messageWidget.insert(tkinter.END, message+"\n")
        else:
            self.messageWidget.insert(tkinter.END, message)
        self.messageWidget.config(state=tkinter.DISABLED)
        self.messageWidget.yview(tkinter.END)
        self.update()

if __name__ == '__main__':
    Gui().mainloop()
