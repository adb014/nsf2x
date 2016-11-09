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

import time
import datetime
import codecs
import os
import ctypes
import subprocess
import traceback
import tempfile
import win32com.client #NB : Calls to COM are starting with an uppercase
import win32com.mapi.mapi
import win32com.mapi.mapitags

try :
    # Python 3.x
    import tkinter
    import winreg
except :
    # Python 2.7
    import Tkinter as tkinter 
    import _winreg as winreg

import mapiex

#FIXME this list should be extended to match regular install paths
notesDllPathList = [r'c:/notes', r'd:/notes', r'c:/program files/notes', r'd:/program files/notes', r'c:/program files (x86)/notes', r'd:/program files (x86)/notes']

def OutlookPath () :
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE")
    n, v, t = winreg.EnumValue(aKey,0)
    winreg.CloseKey(aKey)
    winreg.CloseKey(aReg)
    return v

# The following classes are a means of creating a simple ENUM functionality
# Use list(range()) for Python 2.7 and 3.x compatibility
class Format :
    EML, MBOX, PST = list(range (3))
    
class EncryptionType :
    NONE, ASK, REMOVE, SKIP = list (range (4))
    
class SubdirectoryMBOX :
    NO, YES = list(range (2))

class UsingMAPI :
    NO, YES = list(range (2))
    
class NotesEntries(object) :
    OPEN_RAW_RFC822_TEXT = ctypes.c_uint32(0x01000000)
    OPEN_RAW_MIME_PART = ctypes.c_uint32(0x02000000)
    OPEN_RAW_MIME	= ctypes.c_uint32(0x03000000) # OPEN_RAW_RFC822_TEXT | OPEN_RAW_MIME_PART
    nnotesdll = None
    hDb = ctypes.c_void_p(0)
   
    def __init__(self, fp = None) :
        self.loaddll (fp)
        self.isLoaded(True, False)
        self.SetDLLReturnTypes ()
        stat = self.nnotesdll.NotesInitExtended (0, ctypes.c_void_p(0))
        if (stat != 0) :
            raise NameError("NNOTES DLL can not be initialized (ErrorID %d)" % stat)  
        
        # Throw an error if the DLL didn't load
        self.isLoaded(True, False)
                
    def __delete__ (self) :
        super(NotesEntries, self).__delete__()
        self.NotesTerm ()

    def loaddll (self, fp = None) :
        global nnotesdll
        if (fp != None) :
            if os.path.exists(fp) :
                self.nnotesdll = ctypes.WinDLL(fp)
            else :
                self.nnotesdll = None
        else :
            self.nnotesdll = None
            for p in notesDllPathList :
                fp = os.path.join(p, 'nnotes.dll')
                if os.path.exists(fp) :
                    self.nnotesdll = ctypes.WinDLL(fp)
                    break
      
    def isLoaded (self, raiseError = True, TestDb = True) :
        if raiseError :
            if self.nnotesdll == None :
                raise NameError("NNOTES DLL not loaded")
            elif TestDb and self.hDb == None :
                raise NameError("NNOTES DLL : Database not loaded")
        else :
            return (self.nnotesdll != None and self.hDb != None)

    def SetDLLReturnTypes (self) :
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
        
    def NotesInitExtended (self, argc, argv) :
        self.isLoaded(True, False)
        return self.nnotesdll.NotesInitExtended (argc, argv)
    
    def NotesTerm (self) :
        self.isLoaded (True, False)
        return self.nnotes.NotesTerm ()
    
    def NSFDbOpen (self, path) :
        self.isLoaded(True, False)
        return self.nnotesdll.NSFDbOpen (ctypes.c_char_p(path.encode('utf-8')), ctypes.byref (self.hDb))

    def NSFDbClose (self) :
        self.isLoaded()
        return self.nnotesdll.NSFDbClose (self.hDb)
        
    def NSFNoteCopy (self, hNote) :
        self.isLoaded()
        hNoteNew = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFDbClose (hNote, ctypes.byref (hNoteNew))
        return retval, hNoteNew
        
    def NSFNoteOpenExt (self, nid, flags) :
        self.isLoaded()
        hNote = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFNoteOpenExt (self.hDb, nid, flags, ctypes.byref (hNote))
        return retval, hNote
        
    def NSFNoteOpenByUNID (self, unid, flags) :
        self.isLoaded()
        hNote = ctypes.c_void_p(0)
        retval = self.nnotesdll.NSFNoteOpenByUNID (self.hDb, unid, flags, ctypes.byref (hNote))
        return retval, hNote
        
    def NSFNoteClose (self, hNote) :
        self.isLoaded ()
        return self.nnotesdll.NSFNoteClose (hNote)
        
    def NSFNoteGetInfo (self, hNote, flags) :
        self.isLoaded ()
        retval = ctypes.c_uint16(0)
        self.nnotesdll.NSFNoteGetInfo (hNote, flags, ctypes.byref (retval))
        return retval
        
    def NSFNoteIsSignedOrSealed (self, hNote) :
        self.isLoaded()
        isSigned = ctypes.c_bool(False)
        isSealed = ctypes.c_bool(False)
        retval = self.nnotesdll.NSFNoteIsSignedOrSealed (hNote, ctypes.byref (isSigned), ctypes.byref (isSealed))
        return retval, isSigned.value, isSealed.value
        
    def NSFNoteDecrypt (self, hNote, flags) :
        self.isLoaded()
        return self.nnotesdll.NSFNoteDecrypt (hNote, flags, ctypes.c_void_p(0))

    def NSFItemDelete (self, hNote, iname) :
        self.isLoaded()
        return self.nnotesdll.NSFItemDelete (hNote, iname, len(iname))

    def NSFNoteHasMIMEPart (self, hNote) :
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasMIMEPart (hNote)
   
    def NSFNoteHasMIME (self, hNote) :
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasMIME (hNote)
        
    def NSFNoteHasComposite (self, hNote) :
        self.isLoaded()
        return self.nnotesdll.NSFNoteHasComposite (hNote)
        
    def MMCreateConvControls (self) :
        self.isLoaded()
        hCC = ctypes.c_void_p(0)
        stat = self.nnotesdll.MMCreateConvControls (ctypes.byref (hCC))
        return (stat, hCC)
        
    def MMDestroyConvControls (self, hCC) :
        self.isLoaded()
        return self.nnotesdll.MMDestroyConvControls (hCC)

    def MMSetMessageContentEncoding (self, hCC, flags) :
        self.isLoaded()
        self.nnotesdll.MMSetMessageContentEncoding(hCC, flags)

    def MIMEConvertCDParts (self, hNote, bcanon, bisMime, hCC) :
        self.isLoaded()
        return self.nnotesdll.MIMEConvertCDParts (hNote, bcanon, bisMime, hCC)

    def MIMEConvertMIMEPartsCC (self, hNote, bcanon, hCC) :
        self.isLoaded()
        return self.nnotesdll.MIMEConvertCDParts (hNote, bcanon, hCC)        

    def NSFNoteUpdate (self, hNote, flags) :
        self.isLoaded()
        return self.nnotesdll.NSFNoteUpdate (hNote, flags)
                                
class Gui(tkinter.Frame):
    """Basic Gui for NSF to EML, MBOX, PST export"""
    def __init__(self):
        tkinter.Frame.__init__(self)
        self.master.title("Lotus Notes Converter")
        self.nsfPath = "."
        self.destPath = os.path.join(os.path.expanduser('~'),'Documents')
        self.checked = False
        self.Lotus = None
        self.NotesEntries = None
        self.running = False
        self.dialog = None
        
        # Initialize the default values of the Radio buttons
        self.Format = tkinter.IntVar()
        self.Format.set(Format.EML)
        self.Encrypt = tkinter.IntVar()
        self.Encrypt.set(EncryptionType.SKIP)
        self.MBOXType = tkinter.IntVar()
        self.MBOXType.set(SubdirectoryMBOX.YES)
        self.UseMAPI = tkinter.IntVar()
        self.UseMAPI.set(UsingMAPI.YES)
        
        #Source chooser
        self.chooseNsfButton = tkinter.Button(self.master, text="Select Directory of SOURCE nsf files", command= self.openSource, relief =tkinter.GROOVE, state = tkinter.DISABLED)
        self.chooseNsfButton.grid(row=3,column=1, columnspan=2, sticky=tkinter.E+tkinter.W)

        #Destination chooser
        self.chooseDestButton = tkinter.Button(self.master, text="Select Directory of DESTINATION files", command= self.openDestination, relief =tkinter.GROOVE, state = tkinter.DISABLED)
        self.chooseDestButton.grid(row=3,column=3, columnspan=2, sticky=tkinter.E+tkinter.W)        
        
        #Lotus Password
        tkinter.Label(self.master, text="Enter Lotus Notes password").grid(row=1, column=1, sticky=tkinter.W)
        self.entryPassword = tkinter.Entry(self.master, relief =tkinter.GROOVE) #, show="*")
        self.entryPassword.insert(0, "Enter Lotus Notes password")
        self.entryPassword.grid(row=1,column=1, columnspan=2, sticky=tkinter.E+tkinter.W)
        self.entryPassword.bind("<FocusIn>", self.bindEntry)
      
        #Action button
        self.startButton = tkinter.Button(self.master, text="Open Sessions", command=self.doConvert, relief =tkinter.GROOVE)
        self.startButton.grid(row=1,column=3, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        # Conversion Type
        self.formatTypeEML = tkinter.Radiobutton(self.master, text="EML", variable=self.Format, value=Format.EML)
        self.formatTypeEML.grid(row=2, column=1, sticky=tkinter.E+tkinter.W)
        self.formatTypeMBOX = tkinter.Radiobutton(self.master, text="MBOX", variable=self.Format, value=Format.MBOX)
        self.formatTypeMBOX.grid(row=2, column=2, sticky=tkinter.E+tkinter.W)
        self.formatTypePST = tkinter.Radiobutton(self.master, text="PST", variable=self.Format, value=Format.PST)
        self.formatTypePST.grid(row=2, column=3, sticky=tkinter.E+tkinter.W)
        
        # Options button
        self.optionsButton = tkinter.Button(self.master, text="Options", command=self.doOptions, relief=tkinter.GROOVE, state = tkinter.DISABLED)
        self.optionsButton.grid(row=2,column=4, sticky=tkinter.E+tkinter.W)
        
        #Message Area
        frame = tkinter.Frame(self.master)
        frame.grid(row=4, column=1, columnspan=4)
        self.messageWidget = tkinter.Text(frame, width=80, height=20, state = tkinter.DISABLED, wrap=tkinter.NONE)
        scrollY = tkinter.Scrollbar(frame, orient = tkinter.VERTICAL, command=self.messageWidget.yview)
        self.messageWidget['yscrollcommand'] = scrollY.set
        scrollY.pack(side=tkinter.RIGHT,expand=tkinter.NO,fill=tkinter.Y)
        scrollX = tkinter.Scrollbar(frame, orient = tkinter.HORIZONTAL, command = self.messageWidget.xview)
        self.messageWidget['xscrollcommand'] = scrollX.set
        scrollX.pack(side=tkinter.BOTTOM,expand=tkinter.NO,fill=tkinter.X)

        self.messageWidget.pack(side=tkinter.RIGHT,expand=tkinter.YES,fill=tkinter.BOTH)
        self.log("INFO : Lotus Notes NSF file to EML file converter.")
        self.log("INFO : Contact David.Bateman@edf.fr for more information.\n")
                        
    def openSource(self):
        dirname = self.tk.call('tk_chooseDirectory','-initialdir',self.nsfPath,'-mustexist',True)
        if dirname != "" :
            self.nsfPath = dirname.replace('/','\\')
            self.chooseNsfButton.config(text = "Source directory is : %s" % self.nsfPath)

    def openDestination(self):
        dirname = self.tk.call('tk_chooseDirectory','-initialdir',self.destPath,'-mustexist',True)
        if dirname != "" and type(dirname) is not tuple and str(dirname) != "":
            self.destPath = dirname.replace('/','\\')
            self.chooseDestButton.config(text = "Destination directory is %s" % self.destPath)

    def bindEntry(self, foo= "bar"):
        """Blank the password field and set it in password mode"""
        self.entryPassword.delete(0, tkinter.END)
        self.entryPassword.config(show="*")
        self.entryPassword.unbind("<FocusIn>") #not needed anymore
        self.unchecked()
        
    def check(self):
        if self.Lotus != None :           
            if self.Outlook != None :
                self.checked = True
                self.log("INFO : Connection to Notes and Outlook established")
            else :
                self.unchecked()
                self.log("ERROR : Check that Outlook is running")
        else :
            self.unchecked()
            self.log("ERROR : Check the Notes password")
        return self.checked
        
    def unchecked(self):
        self.startButton.config(text = "Open Sessions")
        self.checked = False
        self.configPasswordEntry()
        
    def configStop(self, AllowButton = True, ActionText = "Stop") :
        self.chooseNsfButton.config(state = tkinter.DISABLED)
        self.chooseDestButton.config(state = tkinter.DISABLED)
        self.entryPassword.config(state = tkinter.DISABLED)
        if AllowButton :
            self.startButton.config(text = ActionText, state = tkinter.NORMAL)
        else :
            self.startButton.config(text = ActionText, state = tkinter.DISABLED)
        self.optionsButton.config(state = tkinter.DISABLED)
        self.formatTypeEML.config(state = tkinter.DISABLED)
        self.formatTypeMBOX.config(state = tkinter.DISABLED)
        self.formatTypePST.config(state = tkinter.DISABLED)

    def configPasswordEntry (self) :
        self.startButton.config(text = "Open Sessions", state = tkinter.NORMAL)
        self.chooseNsfButton.config(text = "Select Directory of SOURCE nsf files", state = tkinter.DISABLED)
        self.chooseDestButton.config(text = "Select Directory of DESTINATION eml files", state = tkinter.DISABLED)
        self.entryPassword.config(state = tkinter.NORMAL)
        self.formatTypeEML.config(state = tkinter.DISABLED)
        self.formatTypeMBOX.config(state = tkinter.DISABLED)
        self.formatTypePST.config(state = tkinter.DISABLED)
        self.optionsButton.config(state = tkinter.DISABLED)

    def configDirectoryEntry (self, SetDefaultPath = True) :
        self.startButton.config(text = "Convert", state = tkinter.NORMAL)
        self.entryPassword.config(state = tkinter.DISABLED)
        self.formatTypeEML.config(state = tkinter.NORMAL)
        self.formatTypeMBOX.config(state = tkinter.NORMAL)
        self.formatTypePST.config(state = tkinter.NORMAL)
        self.optionsButton.config(state = tkinter.NORMAL)

        if SetDefaultPath :
            op = None
            try :
                op = os.path.join(os.path.dirname(self.Lotus.URLDatabase.FilePath),'archive')
            except :
                try :
                    op = os.path.join(os.path.expanduser('~'),'archive') 
                except :
                    op = None
            finally :
                if os.path.exists (op) :
                    self.nsfPath = op
                else :
                    self.nsfPath = '.'
        
            sp = os.path.join(os.path.expanduser('~'),'Documents') 
            if os.path.exists (sp) :
                self.destPath = sp
            else :
                self.destPath = '.'

        # TOBERM
        # This code is just to make my life while testing easier. Remove it eventually
        op = "C:\\Users\\C07056\\Documents\\temp"
        if os.path.exists (op) :
            self.nsfPath = op
            self.destPath = op
        
        self.chooseNsfButton.config(text = "Source directory is : %s" % self.nsfPath)
        self.chooseNsfButton.config(state=tkinter.NORMAL)
        self.chooseDestButton.config(text = "Destination directory is %s" % self.destPath)
        self.chooseDestButton.config(state=tkinter.NORMAL)
        
    def doOptions (self) :
        self.configStop (False, "Convert")
        
        self.dialog = tkinter.Toplevel(master=self.winfo_toplevel())
        self.dialog.title ("NSF2X Options")
        self.dialog.protocol ("WM_DELETE_WINDOW", self.closeOptions)  
        
        L1 = tkinter.Label (self.dialog, text="Use different MBOXes for each sub-folder :")
        L1.grid(row=1, column=1, columnspan=4, sticky=tkinter.W)

        R1 = tkinter.Radiobutton(self.dialog, text="No", variable=self.MBOXType, value=SubdirectoryMBOX.NO)
        R1.grid(row=2, column=1, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        R2 = tkinter.Radiobutton(self.dialog, text="Yes", variable=self.MBOXType, value=SubdirectoryMBOX.YES)
        R2.grid(row=2, column=3, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        L2 = tkinter.Label (self.dialog, text="Treatment of missing encryption certificates for PST conversion :")
        L2.grid(row=3, column=1, columnspan=4, sticky=tkinter.W)
        
        R3 = tkinter.Radiobutton(self.dialog, text="Disable All Encryption", variable=self.Encrypt, value=EncryptionType.NONE)
        R3.grid(row=4, column=1, sticky=tkinter.E+tkinter.W)

        R4 = tkinter.Radiobutton(self.dialog, text="Ask User", variable=self.Encrypt, value=EncryptionType.ASK)
        R4.grid(row=4, column=2, sticky=tkinter.E+tkinter.W)
        
        R5 = tkinter.Radiobutton(self.dialog, text="Remove Recipient", variable=self.Encrypt, value=EncryptionType.REMOVE)
        R5.grid(row=4, column=3, sticky=tkinter.E+tkinter.W)

        R6 = tkinter.Radiobutton(self.dialog, text="Skip Encryption", variable=self.Encrypt, value=EncryptionType.SKIP)
        R6.grid(row=4, column=4, sticky=tkinter.E+tkinter.W)
        
        L3 = tkinter.Label (self.dialog, text="Use MAPI for PST conversion :")
        L3.grid(row=5, column=1, columnspan=4, sticky=tkinter.W)

        R1 = tkinter.Radiobutton(self.dialog, text="No (Old, buggy, slow)", variable=self.UseMAPI, value=UsingMAPI.NO)
        R1.grid(row=6, column=1, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        R2 = tkinter.Radiobutton(self.dialog, text="Yes", variable=self.UseMAPI, value=UsingMAPI.YES)
        R2.grid(row=6, column=3, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        B1 = tkinter.Button(self.dialog, text="Close", command=self.closeOptions, relief=tkinter.GROOVE)
        B1.grid(row=7,column=2, columnspan=2, sticky=tkinter.E+tkinter.W)
        
        self.dialog.focus_force ()
 
    def closeOptions (self) :
        self.configDirectoryEntry(False)
        if self.dialog != None :
            self.dialog.destroy()
        
    def doConvert(self):
        if self.checked:
            if self.running :
                self.running = False;
                self.configStop (False)
                self.log("INFO : Waiting for sub processes to terminate")                
            else :
                self.running = True;                
                self.configStop()
                self.master.after(0, self.doConvertDirectory())
        else : #Check if all is OK
            self.opath = None
            try :
                self.Lotus = win32com.client.Dispatch(r'Lotus.NotesSession')
                if self.NotesEntries == None :
                    self.NotesEntries = NotesEntries()
                # Use rstrip to remove trailing whitespace as not part of the password
                self.Lotus.Initialize(self.entryPassword.get().rstrip())
                self.Lotus.ConvertMime = False
            except Exception as ex:
                self.log("ERROR : Error connecting to Lotus !")
                self.log("ERROR : Exception %s :" % ex)
                # Try to force loading of Notes
                for p in notesDllPathList :
                    fp = os.path.join(p, 'nlsxbe.dll')
                    if os.path.exists(fp) and os.system('regsvr32 /s "%s"' % fp) == 0:
                        break
                self.Lotus = None
                
            try :
                self.Outlook = win32com.client.Dispatch(r'Outlook.Application')
                self.opath = OutlookPath()
                self.log("INFO : Path to Outlook : %s" % self.opath)
            except Exception as ex:
                self.log("ERROR : Could not connect to Outlook !")
                self.log("ERROR : Exception %s :" % ex)
                self.Outlook = None
                
            self.check()
            if self.checked :
                self.configDirectoryEntry()

    def doConvertDirectory(self):
        tl = self.winfo_toplevel()
        self.log("INFO : Starting Convert : %s " % datetime.datetime.now())
        if self.Format.get() == Format.MBOX  and self.MBOXType.get() == SubdirectoryMBOX.NO :
            self.log("WARN : The MBOX file will not have the directory hierarchies present in NSF file")

        for src in os.listdir(self.nsfPath) :
            if not self.running :
                break
        
            abssrc = os.path.join(self.nsfPath, src)         
            if os.path.isfile(abssrc) and src.lower().endswith('.nsf') :
                dest = src[:-4]
                try :
                    self.realConvert(src, dest)
                except Exception as ex:
                    self.log("ERROR : Error converting database %s" % src)
                    self.log("ERROR : Exception %s :" % ex)
                    self.log("ERROR : %s" % traceback.format_exc())
            
        self.log("INFO : End of convert : %s " % datetime.datetime.now())
        tl.title("Lotus Notes Converter")
        self.update()
        self.running = False;
        self.configDirectoryEntry (False)

    def realConvert(self, src, dest):
        """Perform the translation from nsf to X"""
        c = 0 #document counter
        e = 0 #exception counter
        ac = 0 # all message count, though only an upper bounds as some documents not in folders
        tl = self.winfo_toplevel()

        path = os.path.join(self.nsfPath,src)
        try :
            if self.Lotus != None :
                dBNotes = self.Lotus.GetDatabase("", path)
                all = dBNotes.AllDocuments
                ac = all.Count
            else :
                 raise ValueError('Empty Lotus session')       
        except Exception as ex:
            self.log("ERROR : Error connecting to Lotus !")
            self.log("ERROR : Exception %s :" % ex)
            return False

        stat = self.NotesEntries.NSFDbOpen(path)
        if stat != 0 :
            raise ValueError('ERROR : Can not open Lotus database %s with C API (ErrorID %d)' % (path, stat))   
            
        # Open le MBOX
        f = None
        ns = None
        mbox = None
        pst = None
        rootFolder = None

        if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.NO :
            mbox = os.path.join(self.destPath, (dest + ".mbox"))
            self.log("INFO : Opening MBOX file - %s" % mbox)
            f = open (mbox, "wb")
        elif self.Format.get() == Format.PST :
            pst = os.path.join(self.destPath, (dest + ".pst"))
            ns = self.Outlook.GetNamespace(r'MAPI')
            
            self.log("INFO : Opening PST file - %s" % pst)     
            ns.AddStore(pst)
            rootFolder = ns.Folders.GetLast()
            rootFolder.Name = dest
            
            self.log("INFO : Creating directory structure in : %s" % pst)
            all = dBNotes.AllDocuments
            ac = all.Count
            for fld in dBNotes.Views :
                if not self.running :
                    return
                    
                if (fld.Name == "($Sent)" or fld.IsFolder) and fld.EntryCount > 0 :
                    try :
                        pstfld = rootFolder
                        if fld.Name == "($Sent)" :
                            # Special case of Sent folder
                            try :
                                pstfld = pstfld.Folders["Sent"]
                            except :
                                pstfld = pstfld.Folders.Add("Sent")
                                self.log("INFO : Creating Outlook folder %s - Sent" % pst)
                        elif fld.Name == "($Inbox)" : 
                            # Special case of Inbox folder
                            try :
                                pstfld = pstfld.Folders["Inbox"]
                            except :
                                pstfld = pstfld.Folders.Add("Inbox")
                                self.log("INFO : Creating Outlook folder %s - Inbox" % pst)
                        else :
                            for f in fld.Name.split('\\') :
                                try :
                                    pstfld = pstfld.Folders[f]
                                except :
                                    pstfld = pstfld.Folders.Add(f)
                                    self.log("INFO : Creating Outlook folder %s - %s" % (pst, fld.Name))
                    except Exception as ex :
                        self.log("ERROR : Can not create Outlook folder %s - %s" % (pst, fld.Name))
                        self.log("ERROR : %s :" % ex)
                        continue		
                else :
                    continue 
        elif self.Format.get() == Format.EML  :        
            self.log("INFO : Creating directory structure in : %s" % dest)
            all = dBNotes.AllDocuments
            ac = all.Count

            for fld in dBNotes.Views :
                if not self.running :
                    return

                if (fld.Name == "($Sent)" or fld.IsFolder) and fld.EntryCount > 0 :
                    if fld.Name == "($Sent)" :
                        path = os.path.join(self.destPath, dest, "Sent")                    
                    elif fld.Name == "($Inbox)" :
                        path = os.path.join(self.destPath, dest, "Inbox")
                    else :
                        path = os.path.join(self.destPath, dest, fld.Name)
                    try :
                        if not os.path.exists (path) :
                            os.makedirs(path , 0x755)
                            self.log("INFO : Creating directory %s" % path)
                    except Exception as ex :
                        self.log("ERROR : Can not create directory %s" % path)
                        self.log("ERROR : %s :" % ex)
                        continue                
                else :
                    continue
                    
        # Preconvert all messages to MIME before writing EML files as the
        # C DLL might not be finished saving the message before the COM
        # interface tries to access the MIME body. Also the call to mapiex.mapi()
        # must come after the conversion, as if it doesn't the call to
        # MIMEConvertCDParts will raise a "File does not exist error (259)".
        # ?*#! -> Weird interaction MAPI to Notes  
        self.log("INFO : Starting MIME encoding of messages")
        for fld in dBNotes.Views :
            if  not (fld.Name == "($Sent)" or fld.IsFolder) or fld.EntryCount <= 0 :
                if fld.EntryCount > 0 :
                    tl.title("Lotus Notes Converter - Phase 2/3 Converting MIME (%.1f%%)" % float(10.*c/ac))
                    self.update()
                continue
            doc = fld.GetFirstDocument()
            while doc and e < 100 : #stop after 100 exceptions...
                if not self.running :
                    return
                    
                try :              
                    if not self.ConvertToMIME(doc) :
                        e+=1
                        self.log("ERROR : Can not convert message %d to MIME" % c)
                except Exception as ex:
                    self.log("ERROR : Exception converting message %d to MIME : %s" % (c, ex))
                doc = fld.GetNextDocument(doc)
                c+=1
                if (c % 20) == 0:
                    tl.title("Lotus Notes Converter - Phase 2/3 Converting MIME (%.1f%%)" % float(10.*c/ac))
                    self.update()

        if e == 100 :
            self.log ("ERROR : Too many exceptions. Returning")
                    
        MAPI = None
        if self.UseMAPI.get() == UsingMAPI.YES :
            try :
                MAPI = mapiex.mapi()        
                MAPI.OpenMessageStore(dest)
                MAPIrootFolder = MAPI.OpenRootFolder ()
            except Exception as ex:
                self.log("ERROR : Could not connect to MAPI !")
                self.log("ERROR : Exception %s :" % ex)
                raise
                
        self.log("INFO : Starting importation of EML messages into mailbox")
        ac = c # Update all message count
        c=0
        e=0
        for fld in dBNotes.Views :
            if  not (fld.Name == "($Sent)" or fld.IsFolder) or fld.EntryCount <= 0 :
                if fld.EntryCount > 0 :
                    tl.title("Lotus Notes Converter - Phase 3/3 Import Message %d of %d (%.1f%%)" % (c, ac, float(10.*(ac + 9.*c)/ac)))
                    self.update()
                continue

            pstfld = None
            if self.Format.get() == Format.PST :
                if self.UseMAPI.get() == UsingMAPI.YES :
                    if fld.Name == "($Sent)" :
                        pstfld = MAPIrootFolder.OpenSubFolder("Sent")
                    elif fld.Name == "($Inbox)" :
                        pstfld = MAPIrootFolder.OpenSubFolder("Inbox")
                    else :
                        pstfld = MAPIrootFolder.OpenSubFolder (fld.Name)
                else :
                    pstfld = rootFolder
                    if fld.Name == "($Sent)" :
                        pstfld = pstfld.Folders["Sent"]
                    elif fld.Name == "($Inbox)" :
                        pstfld = pstfld.Folders["Inbox"]
                    else :
                        for f in fld.Name.split('\\') :
                            pstfld = pstfld.Folders[f]                        
            elif self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.YES :
                mbox = None
                if fld.Name == "($Sent)" :
                    mbox = os.path.join(self.destPath, dest, "Sent.mbox")
                elif fld.Name == "($Inbox)" :
                    mbox = os.path.join(self.destPath, dest, "Inbox.mbox")
                else :
                    mbox = os.path.join(self.destPath, dest, (fld.Name + ".mbox"))                   

                try :
                    mboxdir = os.path.dirname (mbox)
                    if not os.path.exists (mboxdir) :
                        os.makedirs(mboxdir, 0x755)
                        self.log("INFO : Creating directory %s" % mboxdir)
                except Exception as ex :
                    self.log("ERROR : Can not create directory %s" % mboxdir)
                    self.log("ERROR : %s :" % ex)
                
                self.log("INFO : Opening MBOX file - %s" % mbox)
                f = open (mbox, "wb")
                
            doc = fld.GetFirstDocument()
            d=1
            while doc and e < 100 : #stop after 100 exceptions...
                if not self.running :
                    return
                    
                try :
                    eml = None
                    
                    if doc.GetMIMEEntity("Body") == None :
                        subject = doc.GetFirstItem("Subject")
                        self.log("WARN : Message %d has no MIME body" % c)
                        if subject :
                            self.log ("     : Subject : %s" % subject.Text)
                        self.log ("     : Skipping as probably not a message")
                    else :                
                        if self.Format.get() != Format.MBOX :
                            if self.Format.get() == Format.EML :
                                if fld.Name == "($Sent)" :
                                    eml = os.path.join(self.destPath, dest, "Sent", (str(d) + ".eml"))
                                elif fld.Name == "($Inbox)" :
                                    eml = os.path.join(self.destPath, dest, "Inbox", (str(d) + ".eml"))
                                else :
                                    eml = os.path.join(self.destPath, dest, fld.Name, (str(d) + ".eml"))
                                f = open (eml, "wb")  # Need to treat as binary so that windows doesn't convert \n\r to \n\n\r    
                            elif self.Format.get() == Format.PST :
                                (fd, eml) = tempfile.mkstemp(suffix=".eml")
                                f = os.fdopen (fd, "wb")
                            
                        if  self.WriteMIMEOutput (f, doc) :
                            d+=1
                            if self.Format.get() == Format.PST :                            
                                f.close ()
                                self.ConvertEMLToOutlook (doc, eml, ns, c, pstfld)

                                # Done with the temporary EML file. Remove it
                                if eml != None :
                                    os.remove (eml)                            
                                
                            elif self.Format.get () == Format.EML :
                                f.close ()
                        else :
                            raise NameError("Can not write Lotus MIME message to a file")
                        
                except Exception as ex:
                    e += 1 #count the exceptions
                    if self.Format.get () != Format.MBOX :
                        # File might already be closed and/or removed. So failure is ok
                        try:
                            f.close ()
                        except :
                            pass
                        try :
                            os.remove(eml)
                        except :
                            pass
                    self.log("ERROR : Exception for message %d (%s) :" % (c, ex))
                    self.log("ERROR : %s" % traceback.format_exc())
        
                finally:                  
                    c+=1
                    doc = fld.GetNextDocument(doc)
                    
                    if self.Format.get() == Format.MBOX :
                        # MBOX is recognized by "\nFrom:" string. So add a trailing \n to each message to ensure this format
                        f.write(b"\n")
 
                    if (c % 20) == 0:
                        tl.title("Lotus Notes Converter - Phase 2/3 Import Message %d of %d (%.1f%%)" % (c, ac, float(10.*(ac + 9.*c)/ac)))
                        self.update()
                       
            if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.YES :
                f.close ()

        if self.Format.get() == Format.MBOX and self.MBOXType.get() == SubdirectoryMBOX.NO :
            f.close ()
        self.log("\nINFO : Finished populating directory : %s" % dest)
        self.log("INFO : Exceptions to treat manually: %d ... Documents OK : %d Untreated Documents : %d" % (e, c - e, ac - c))

        return True
        
    def ConvertEMLToOutlook (self, doc, eml, ns, id, pstfld) :
    
        if self.UseMAPI.get() == UsingMAPI.YES :
            message = pstfld.ImportEML(eml)

            enc = doc.GetFirstItem("Encrypt")
            if enc != None and enc.Text == '1' :
                # Reopen as a MailItem and then encrypt. It would be better
                # to work directly with the IMessage though that seems rather
                # involved. See the site
                # https://blogs.msdn.microsoft.com/webdav_101/2015/12/16/about-encrypting-or-signing-a-message-programmatically/
                # for information on how to do this
                entryID = message.GetEntryID()
                m = ns.GetItemFromID (codecs.encode(entryID[1][0][1], 'hex'))

                if m == None :
                    # Got nothing
                    raise NameError("Can not open MailItem from MAPIMessage (message %d)" % id)
                try :
                    self.CheckAndEncrypt (ns, id, m)
                except Exception as ex :
                    self.log ("WARN : Can not encrypt message %d (%s)" % (id, ex))
                m.Close(0)

            del message   # Explicitly delete the message so IUnknown:Release is called
        else :    
            # Load the EML file into the Outlook UI.
            try :          
                subprocess.call([self.opath, "/eml", eml])
            except OSError as e:
                if e.errno == errno.EACCES :
                    # There is an occasional race every 1000
                    # messages or so. Just wait for the OS to
                    # properly close the EML file
                    time.sleep(0.05)
                    subprocess.call([self.opath, "/eml", eml])
                else :
                    raise                                      
            
            # Load the list of all the open Outlook inspectors and check for the
            # one we are interested. It is identified by its Sender and Date.
            # Don't use ActiveInspector as dialogs and popups can confuse the
            # issue. It should also allow the screen to be locked during the
            # importation.
            # If the Message-ID field exists rely on it. Otherwise, don't use 
            # the Subject field as it seems the the text encoding between
            # Outlook and Notes can make some messages be falsely identified as
            # different

            m1 = None
            retries = 0   
            nSender = None
            nDate = None
            nID = None

            # If we have a Message_ID we can rely on it to identify the mail
            tmp = doc.GetFirstItem("$MessageID")
            if tmp == None :     
                for notesSenderField in ("Sender", "From", "Principal", "InetFrom") :
                    tmp = doc.GetFirstItem(notesSenderField)
                    if tmp != None :
                        nSender = tmp.Text
                        break
                if nSender == None :
                    self.log("ERROR : Can't get message %d sender address" % id)
                else :
                    nDate = doc.GetFirstItem("PostedDate")
                    if nDate == None :
                        self.log("ERROR : Can't get message %d sent date" % id)
                    else :
                        try :
                            nDate = int(time.mktime(time.strptime(nDate.Text, "%d/%m/%Y %H:%M:%S")))
                        except :
                            # Can't rely on the date from Outlook in this case for the test below.
                            nDate = int(time.mktime(time.strptime(nDate.Text, "%d/%m/%Y")))
            else :
                nID = tmp.Text
            
            while (nID != None or (nSender != None and nDate != None)) and retries < 100 :
                try :
                    # Don't use "for inspector in self.Outlook.Inspectors :" as doesn't seem
                    # to work with older versions of the OS.
                    
                    for i in list (range (1, len(self.Outlook.Inspectors) + 1)) :               
                        m2 = win32com.client.CastTo (self.Outlook.Inspectors.Item(i).CurrentItem, "_MailItem")

                        if nID != None :
                            if nID != m2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") :
                                continue
                        else :
                            oSender = m2.SenderEmailAddress
                            if not oSender or oSender != nSender :
                                continue         
                            oDate = m2.SentOn
                            if not oDate or nDate != int(oDate) :
                                continue              
                         
                        m1 = m2
                        break                                 
                    if m1 != None :
                        break
                    else :
                        time.sleep (0.05) # Sleep for 50 ms
                except :
                    time.sleep (0.05) # Sleep for 50 ms
                finally :                                        
                    retries += 1
                
            if m1 == None :
                # Got nothing after 5 seconds
                raise NameError("Can not open EML message %d with Outlook" % id)
                
            enc = doc.GetFirstItem("Encrypt")
            if enc != None and enc.Text == '1' :
                try :
                    self.CheckAndEncrypt (ns, id, m1)
                except Exception as ex :
                    self.log ("WARN : Can not encrypt message %d (%s)" % (id, ex))

            m1.Move(pstfld)
            m1.Close(1)     # Discard the copy in the Outlook Inbox
        
    def CheckAndEncrypt (self, ns, id, m) :
        if self.Encrypt.get() == EncryptionType.NONE :
            return

        if self.Encrypt.get() != EncryptionType.ASK :            
            # Check if can resolve the sender
            r = ns.CreateRecipient (m.SenderEmailAddress)
            r.Resolve()                
            try :
                r.AddressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x8C6A1102")
            except : 
                try :
                    # The certificate might be in a contact 
                    contact = r.AddressEntry.GetContact()
                    contact.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A701102")
                except :            
                    self.log ("WARN : Message %d not encrypted !!!! Sender %s has missing certificate" % (id, m.SenderEmailAddress)) 
                    # Outlook 2007 doesn't have an m.Sender attribute of type AddressEntry and so can't easily
                    # do anything else but not encrypt the mail. Seem to exist in Outlook 2010/2013 however
                    return         
        
            # Check the recipients addresses can be resolved and we have a valid certificate for them
            for i in range (m.Recipients.Count, 0, -1) :
                r = ns.CreateRecipient (m.Recipients[i].Name)
                r.Resolve ()
                try :
                    r.AddressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x8C6A1102")
                except :
                    try :
                        contact = r.AddressEntry.GetContact()
                        contact.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A701102")
                    except :
                        if self.Encrypt.get() == EncryptionType.REMOVE :
                            self.log ("WARN : Removing %s from recipients of message %d" % (m.Recipients[i].Name, id))
                            m.Recipients.Remove(i)
                        else :
                            self.log ("WARN : Can not encrypt to %s in message %d" % (m.Recipients[i].Name, id))
                            return
                        
        # Ok now we can flag the mail as encrypted. 
        m.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x6E010003", 1)
 
    def ConvertToMIME (self, doc) :
        kprivate = "$KeepPrivate";
        # I'd really like to use doc.UniversalID here to open the file with 
        # NSFNoteOpenByUNID. However, doc.UniversalID is a string and
        # NSFNoteOpenByUNID expects a struct and the conversion between the
        # two doesn't seem easy. Use doc.NoteID instead
        # stat, hNote = self.NotesEntries.NSFNoteOpenByUNID(doc.UniversalID, self.NotesEntries.OPEN_RAW_MIME)
        stat, hNote = self.NotesEntries.NSFNoteOpenExt(ctypes.c_uint32(int(doc.NoteID, 16)), self.NotesEntries.OPEN_RAW_MIME)

        if stat != 0 :
             self.log ("ERROR : Can not open document id 0x%s (ErrorID : %d)" % (doc.NoteID, stat))
        else :
            try :
                # The C API identifies some unencrypted mail as "Sealed". These don't need
                # to be unencrypted to allow conversion to MIME.
                enc = doc.GetFirstItem("Encrypt")
                if enc != None and enc.Text == '1' : 
                    # if the note is encrypted, try to decrypt it. If that fails
                    #(e.g., we don't have the key), then we can't convert to MIME
                    # (we don't care about the signature)
                    retval, isSigned, isSealed = self.NotesEntries.NSFNoteIsSignedOrSealed(hNote)
                    if isSealed :
                        # self.log ("INFO : Document note id 0x%s is encrypted." % doc.NoteID)
                        DECRYPT_ATTACHMENTS_IN_PLACE = ctypes.c_uint16(1);
                        stat = self.NotesEntries.NSFNoteDecrypt(hNote, DECRYPT_ATTACHMENTS_IN_PLACE);
                        
                        if stat != 0 :
                            self.log ("ERROR : Document note id 0x%s is encrypted, cannot be converted." % doc.NoteID)
                
                if stat == 0 :
                    # If present, $KeepPrivate will prevent conversion, so nuke the sucka
                    self.NotesEntries.NSFItemDelete(hNote, kprivate);

                    # if the note is already in mime format, we don't have to convert
                    if (self.NotesEntries.NSFNoteHasComposite(hNote)) :
                        stat, hCC = self.NotesEntries.MMCreateConvControls ()
                        if stat == 0 :
                            self.NotesEntries.MMSetMessageContentEncoding(hCC, 2) # html w/images & attachments
                            
                            # NOTE_FLAG_CANONICAL = 0x4000 see nsfnote.h
                            _NOTE_FLAGS = ctypes.c_uint16 (7)
                            bCanonical = (self.NotesEntries.NSFNoteGetInfo (hNote, _NOTE_FLAGS).value) & 0x4000 != 0
                            bIsMime = self.NotesEntries.NSFNoteHasMIMEPart(hNote)
                            stat = self.NotesEntries.MIMEConvertCDParts(hNote, bCanonical, bIsMime, hCC)
                            
                            if stat == 0 :
                                UPDATE_FORCE = ctypes.c_uint16(1);
                                stat = self.NotesEntries.NSFNoteUpdate(hNote, UPDATE_FORCE)
                                if stat != 0 :
                                    self.log("ERROR : Error calling NSFNoteUpdate (%d)" % stat)
                            elif stat == 14941 :
                                self.log("INFO : MIMEConvertCDParts : Error converting note id 0x%s to MIME type text/html" % doc.NoteID)
                                self.log("INFO : MIMEConvertCDParts : Attempting to convert to text/plain")
                                self.NotesEntries.MMSetMessageContentEncoding(hCC, 1)
                                stat = self.NotesEntries.MIMEConvertCDParts(hNote, bCanonical, bIsMime, hCC)    
                                
                            if stat != 0 :
                                self.log ("ERROR : Error calling MIMEConvertCDParts (%d)" % stat)
                                
                            self.NotesEntries.MMDestroyConvControls(hCC)
                        else :
                            self.log("ERROR : Error calling MMCreateConvControls (%d)" % stat)
                            
                if hNote != None :
                    self.NotesEntries.NSFNoteClose(hNote)
            except :
                if hNote != None :
                    # Ensure Note is closed and then re-raise the exception
                    self.NotesEntries.NSFNoteClose(hNote)
                raise
            
        return (stat == 0)   
        
    def WriteMIMEChildren (self, f, mime, first) :
        if mime != None :
            contentType = mime.ContentType;
            headers = mime.Headers;
            encoding = mime.Encoding;
            
            # if it's a binary part, force it to b64
            if (encoding == 1730 or encoding == 1729) :  
                # MIMEEntity.ENC_IDENTITY_BINARY and MIMEEntity.ENC_IDENTITY_8BIT
                mime.EncodeContent(1727)  # MIMEEntity.ENC_BASE64
                headers = mime.Headers

            if first :
                # Place the From and Date fields first to simplify conversion to MBOX format
                content = mime.GetSomeHeaders(["From"], True)
                f.write(content.encode('utf-8'))
                if not content.endswith ("\n") :
                    f.write (b"\n")
                content = mime.GetSomeHeaders(["Date"], True)
                f.write(content.encode('utf-8'))
                if not content.endswith ("\n") :
                    f.write (b"\n")
                
                # message envelope. If no MIME-Version header, add one
                if "MIME-Version:" not in headers :
                    f.write(b"MIME-Version: 1.0\n")
                
                # Write the rest of the headers, but exclude the MIME content-type to be placed last
                content = mime.GetSomeHeaders(["From", "Date", "Content-type"], False)
                # Some of the text might be in utf-8 so give it special treatment
                f.write(content.encode('utf-8'))
                if not content.endswith ("\n") :
                    f.write (b"\n")
                    
                content = mime.GetSomeHeaders(["Content-type"], True)
                f.write(content.encode('utf-8'))
                if not content.endswith ("\n") :
                    f.write (b"\n")
            else :
                f.write (headers.encode('utf-8'))
                if not headers.endswith ("\n") :
                    f.write (b"\n")

            f.write(b"\n")       
            content = mime.ContentAsText
            f.write (content.encode('utf-8'))
            if not content.endswith ("\n") :
                f.write (b"\n")
                    
            f.flush ()       
                    
            if (contentType.startswith("multipart")) :
                content = mime.preamble
                if (content != "") :
                    f.write (content.encode('utf-8'))
                    if not content.endswith("\n") :
                        f.write (b"\n")
                                                
                child = mime.GetFirstChildEntity ()
                while child != None :
                    content = child.BoundaryStart
                    f.write (content.encode('utf-8'))
                    if not content.endswith("\n") :
                        f.write (b"\n")

                    self.WriteMIMEChildren (f, child, False)
                        
                    content = child.BoundaryEnd
                    f.write (content.encode('utf-8'))
                    if not content.endswith("\n") :
                        f.write (b"\n")                   
                        
                    child = child.GetNextSibling ()

    def WriteMIMEOutput (self, f, doc) :
        if doc != None :
            # Get first Body item with a MIME encoding
            mE = doc.GetMIMEEntity("Body")
            if mE != None :
                self.WriteMIMEChildren (f, mE, True)
                return True
            else :
                self.log("WARN : Message 0x%s has no MIME body" % doc.NoteID)
                self.log("      Type : %d" % doc.GetFirstItem("Body").Type)
                self.log("      Subject : %s" % doc.GetFirstItem("Subject").Text)
        return False

    def log(self, message = "", newline = True):
        self.messageWidget.config(state = tkinter.NORMAL)
        if (newline) :
            self.messageWidget.insert(tkinter.END, message+"\n")
        else :
            self.messageWidget.insert(tkinter.END, message)
        self.messageWidget.config(state = tkinter.DISABLED)
        self.messageWidget.yview(tkinter.END)
        self.update()                
                
if __name__ == '__main__':
    Gui().mainloop()
