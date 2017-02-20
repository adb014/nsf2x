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

# A minimalist class wrapping the MAPI functionality that is needed. Look
# at the code at  http://www.codeproject.com/internet/CMapiEx.asp for ideas
# of how to extend this code if needed

import os
import win32com.mapi.mapi
import win32com.mapi.mapitags
from win32com.client import pythoncom
from win32com.server import util
import pywintypes

class mapiobject (object) :
    def __init__ (self, mapi, item = None) :
        self.mapi = mapi
        self.item = item
        
    def GetProperty (self, prop) :
        return self.item.GetProps([prop])
        
    def GetPropertyValue (self, prop) :
        p = self.item.GetProps([prop])
        return p.value
        
    def SetProperty (self, prop, value) :
        self.item.SetProps([(prop, value)])

    def GetEntryID (self) :
        return self.GetProperty(win32com.mapi.mapitags.PR_ENTRYID)
        
    def Save (self, flags = 0) :
        self.item.SaveChanges(flags)
        
    def Open (self, eid) :
        self.mapi.session().OpenEntry (eid, None, win32com.mapi.mapi.MAPI_BEST_ACCESS)
               
class mapimessage (mapiobject) :
    MSGFLAG_READ = 1
    MSGFLAG_UNMODIFIED = 2
    MSGFLAG_UNSENT = 4
    MSGFLAG_FROMME = 8
    
    def __init__ (self, mapi, m = None) :
        super(mapimessage, self).__init__(mapi, m)
        
    def message (self) :
        return self.item

    def GetSubject (self) :
        return self.GetProperty (win32com.mapi.mapitags.PR_SUBJECT)

    def SetSubject (self, subject) :
        return self.SetProperty (win32com.mapi.mapitags.PR_SUBJECT, subject)
        
    def GetBody (self) :
        return self.GetProperty (win32com.mapi.mapitags.PR_BODY)
        
    def SetBody (self, body) :
        return self.SetProperty (win32com.mapi.mapitags.PR_BODY, body)
        
    def GetMessageFlags (self) :
        return self.GetProperty (win32com.mapi.mapitags.PR_MESSAGE_FLAGS)
    
    def SetMessageFlags (self, flags) :
        return self.SetProperty (win32com.mapi.mapitags.PR_MESSAGE_FLAGS, flags)
        
    def ImportEML (self, eml) :            
        f = open(eml, "rb")   
        self.mapi.MimeToMapi(f, self.item, 0x20)
        self.SetMessageFlags(self.MSGFLAG_READ)
        f.close()
        
class mapiappointment (mapiobject) :
    OUTLOOK_DATA2 = 0x00062002
    OUTLOOK_APPOINTMENT_START = 0x820D
    OUTLOOK_APPOINTMENT_END = 0x820E
    OUTLOOK_APPOINTMENT_LOCATION = 0x8208

    def __init__ (self, mapi, a = None) :
        super(mapimessage, self).__init__(mapi, a)
        
    def appointment (self) :
        return self.item
        
    def GetSubject (self) :
        return self.GetProperty (win32com.mapi.mapitags.PR_SUBJECT)

    def SetSubject (self, subject) :
        return self.SetProperty (win32com.mapi.mapitags.PR_SUBJECT, subject)
        
    def GetBody (self) :
        return self.GetProperty (win32com.mapi.mapitags.PR_BODY)
        
    def SetBody (self, body) :
        return self.SetProperty (win32com.mapi.mapitags.PR_BODY, body)

class mapifolder (mapiobject) :
    def __init__ (self, mapi, f, n = None) :
        super(mapifolder, self).__init__ (mapi, f)
        self.name = n
        self.Hierarchy = None
        self.contents = None
        
    def folder (self) :
        return self.item

    def GetHierarchy (self) :
        self.Hierarchy = self.folder().GetHierarchyTable(0)
        if self.Hierarchy != None :
            self.Hierarchy.SetColumns((win32com.mapi.mapitags.PR_ENTRYID, win32com.mapi.mapitags.PR_DISPLAY_NAME),0)
        return self.Hierarchy
        
    def _splitpath (self, path) :
        flds = []
        while True :
            [path, tail] = os.path.split (path)
            if len(tail) == 0 :
                break
            flds.insert (0, tail)
        return flds
    
    def OpenSubFolder (self, flds) :
        self.GetHierarchy()
  
        if not isinstance (flds, list) :
            flds = self._splitpath (flds)

        fld = flds[0]
        flds = flds[1:]
  
        while True :
            subfolder = self.GetNextSubFolder()
            if subfolder == None:
                return None
            elif subfolder.name == fld :
                if len (flds) == 0 :
                    return subfolder
                else :
                    return subfolder.OpenSubFolder (flds)

    def CreateSubFolder (self, flds) :
        if not isinstance (flds, list) :
            flds = self._splitpath (flds)

        fld = flds[0]
        flds = flds[1:]
        subfolder = mapifolder(self.mapi, self.folder().CreateFolder(win32com.mapi.mapi.FOLDER_GENERIC, fld, None, None, win32com.mapi.mapi.OPEN_IF_EXISTS | win32com.mapi.mapi.MAPI_UNICODE)) 
        if subfolder == None or len (flds) == 0 :
            return subfolder
        else :
            return subfolder.CreateSubFolder (flds)

    def GetFirstSubFolder (self) :
        if self.GetHierarchy() :
            return self.GetNextSubFolder ()
        else :
            return None

    def GetNextSubFolder (self) :
        if self.Hierarchy == None :
            raise "mapifolder:GetNextSubFolder : Call GetFirstSubFolder before GetNextSubFolder"
        subfolder = None
        while True :
            rows = self.Hierarchy.QueryRows(1, 0)
            #if this is the last row then stop
            if len(rows) != 1:
                break
            row = rows[0]
            (eid_tag, eid), (name_tag, name) = row
            subfolder = self.folder().OpenEntry(eid, None, win32com.mapi.mapi.MAPI_MODIFY)
            if subfolder != None :
                subfolder = mapifolder (self.mapi, subfolder, name)
                break
        return subfolder
    
    def CreateMessage (self) :
        return mapimessage (self.mapi, self.folder().CreateMessage (None, 0))
        
    def GetContents (self) :
        try :
            self.contents = self.folder().GetContentsTable(0)
            self.contents.SetColumns((win32com.mapi.mapitags.PR_ENTRYID, 
                                win32com.mapi.mapitags.PR_MESSAGE_FLAGS),0)
        except :
            self.contents = None
        return self.contents
 
    def GetFirstMessage (self) :
        self.GetContents ()
        return self.GetNextMessage ()
        
    def GetNextMessage (self) :
        if self.contents == None :
            raise "mapifolder:GetNextMessage : Call GetFirstMessage before GetNextMessage"
            
        rows = self.contents.QueryRows(1, 0)
        if len(rows) != 1:
            return None
        row = rows[0] 
        (eid_tag, eid), (flag_tag, flag) = row
        message = mapimessage(self.mapi)
        message.Open(eid)
        return message
    
    def GetFirstAppointment (self) :
        self.GetContents ()
        return self.GetNextAppointment ()
        
    def GetNextAppointment (self) :
        if self.contents == None :
            raise "mapifolder:GetNextAppointment : Call GetFirstAppointment before GetNextAppointment"
            
        rows = self.contents.QueryRows(1, 0)
        if len(rows) != 1:
            return None
        row = rows[0] 
        (eid_tag, eid), (flag_tag, flag) = row
        appointment = mapiappointment(self.mapi)
        appointment.Open(eid)
        return appointment
        
    def ImportEML (self, eml) :                  
        message = self.CreateMessage()
        message.ImportEML(eml)
        message.Save()
        return message
        
class mapi (object) :
    def __init__ (self, profilename = "") :
        # FIXME
        # The MAPI initialisation changes the directory. Something that can
        # mess with the paths to other files. Save the path before initialisation
        # and restore it
        save_cwd = os.getcwd()
        win32com.mapi.mapi.MAPIInitialize(None)
        self.messagestorestable = None
        self.converter = None
        self._session = win32com.mapi.mapi.MAPILogonEx(0, profilename, None, win32com.mapi.mapi.MAPI_EXTENDED | win32com.mapi.mapi.MAPI_USE_DEFAULT)
        os.chdir(save_cwd)
   
    def __delete__ (self) :
        win32com.mapi.mapi.MAPIUninitialize()
           
    def session (self) :
        return self._session
        
    def GetProfileName (self) :
        StatusTable = self.session().GetStatusTable(0)
        StatusTable.SetColumns((win32com.mapi.mapitags.PR_DISPLAY_NAME_A, win32com.mapi.mapitags.PR_RESOURCE_TYPE),0)       
        while True :
            rows = StatusTable.QueryRows(1, 0)
            #if this is the last row then stop
            if len(rows) != 1:
                break
            row = rows[0]
            (name_tag, name), (res_tag, res) = row
            if res == 39 :   # MAPI_SUBSYSTEM = 39
                return name
        return ""
        
    def GetProfileEmail (self) :
        try :
            eid = self.session().QueryIdentity()
            AddressBook = self.session().OpenAddressBook(0, None, win32com.mapi.mapi.AB_NO_DIALOG)
            obj = AddressBook.OpenEntry((eid), None, win32com.mapi.mapi.MAPI_BEST_ACCESS)
            try :
                # FIXME : Why is PR_SMTP_ADDRESS not in win32com.mapi.mapitags ? 
                PR_SMTP_ADDRESS = int(0x39FE001F)
                (count, prop) = obj.GetProps ((PR_SMTP_ADDRESS), 0)
                return prop[0][1]
            except :
                (count, prop) = obj.GetProps ((win32com.mapi.mapitags.PR_EMAIL_ADDRESS), 0)
                return prop[0][1]
        except Exception as ex:
            pass
        return ""               
        
    def MimeToMapi (self, mimestream, m, flag = 0) :
        if self.converter == None :
            ## CLSID_IConverterSession
            clsid = pywintypes.IID('{4e3a7680-b77a-11d0-9da5-00c04fd65685}')
            ## IID_IConverterSession
            iid = pywintypes.IID('{4b401570-b77b-11d0-9da5-00c04fd65685}') 
            
            tmp = pythoncom.CoCreateInstance (clsid, None, pythoncom.CLSCTX_INPROC_SERVER, pythoncom.IID_IUnknown)
            self.converter = tmp.QueryInterface (iid)

        Istrm = util.wrap (util.FileStream(mimestream), pythoncom.IID_IStream)
        self.converter.MIMEToMAPI(Istrm, m, flag)        
        
    def _GetContents (self) :
        self.messagestorestable = self.session().GetMsgStoresTable(0)
        self.messagestorestable.SetColumns((win32com.mapi.mapitags.PR_ENTRYID, 
                                    win32com.mapi.mapitags.PR_DISPLAY_NAME_A, 
                                    win32com.mapi.mapitags.PR_DEFAULT_STORE),0)
        return self.messagestorestable

    def GetMessageStoreNames (self) :
        Names = []
        self._GetContents ()
        while True:
            rows = self.messagestorestable.QueryRows(1, 0)
            #if this is the last row then stop
            if len(rows) != 1:
                break
            row = rows[0]
            # unpack the row and print name of the message store
            (eid_tag, eid), (name_tag, name), (def_store_tag, def_store) = row
            Names.append (name)           
        return Names
    
    def OpenMessageStore (self, storename = None) :
        self.msgstore = None
        row = None
        self._GetContents ()
        while True:
            rows = self.messagestorestable.QueryRows(1, 0)
            #if this is the last row then stop
            if len(rows) != 1:
                raise "mapi:OpenMessageStore : Error opening message store"
            #if this store has the right name stop
            if storename == None or ((win32com.mapi.mapitags.PR_DISPLAY_NAME_A,storename.encode('utf-8')) in rows[0]):
                row = rows[0]
                break

        if row == None :
            if storename == None :
                raise "mapi:OpenMessageStore : Can not find default messagestore"
            else :
                raise "mapi:OpenMessageStore : Can not find messagestore : %s" % storename

        # unpack the row and open the message store
        (eid_tag, eid), (name_tag, name), (def_store_tag, def_store) = row
        self.msgstore = self.session().OpenMsgStore(0, eid, None, win32com.mapi.mapi.MDB_NO_DIALOG | win32com.mapi.mapi.MAPI_BEST_ACCESS)
        self.storename = storename
        
    def AddMessageStore (self, storename, storepath) :
        # Note that this method adds the store without opening it.
        if not os.path.exists (storepath) :
            raise NameError("mapi:AddMessageStore : File %s does not exist" % storepath)
        
        # FIXME: would like to use self.session.AdminServices(0), rather than
        # the next lines, but not in PyWin32 yet, so have to get the current
        # profile from the ProfileTable

        # Identify current profile
        profileAdmin = win32com.mapi.mapi.MAPIAdminProfiles(0)
        profileTable = profileAdmin.GetProfileTable(0)
        profileRows = win32com.mapi.mapi.HrQueryAllRows(profileTable, [win32com.mapi.mapitags.PR_DISPLAY_NAME_A], None, None, 0)
        profilename = self.GetProfileName()
        profile = None
        for p in profileRows :
            if p[0][1] == profilename :
                profile = p[0][1]
                break
        if not profile :
            raise NameError("mapi:AddMessageStore : Can not identify profile %s" % profilename)
          
        # Add the PST file as a service to the profile. If 'MSPST MS' service is not defined in
        # the file mapisvc.inf then this will fail with MAPI_E_NOT_FOUND. This seems to be the case
        # in most recent windows installs !!!
        serviceAdmin = profileAdmin.AdminServices(str(profile, 'utf-8'), None, 0, win32com.mapi.mapi.MAPI_UNICODE) 
        serviceAdmin.CreateMsgService('MSPST MS', None, 0, win32com.mapi.mapi.MAPI_UNICODE)
        
        # Get the service table - looking for service IDs. The PST is the last
        # service added
        # FIXME : Is there a race condition here that could be a security risk ?
        msgServiceTable = serviceAdmin.GetMsgServiceTable(0)
        msgServiceRows = win32com.mapi.mapi.HrQueryAllRows(msgServiceTable, [win32com.mapi.mapitags.PR_SERVICE_UID], None, None, 0)
        serviceUID = msgServiceRows[-1][0][1]
        serviceUID = pythoncom.MakeIID(serviceUID, 1)
        
        # Configure the PST file.
        # FIXME : Why is PR_PST_PATH not in win32com.mapi.mapitags ? 
        PR_PST_PATH = int(0x6700001E)
        serviceAdmin.ConfigureMsgService(serviceUID, 0, 0, ((win32com.mapi.mapi.PR_DISPLAY_NAME_A, storename), (PR_PST_PATH, storepath)))
    
    def OpenRootFolder (self) :
        # Open the root folder of the MsgStore 
        hr, props = self.msgstore.GetProps((win32com.mapi.mapitags.PR_IPM_SUBTREE_ENTRYID), 0)
        (tag, eid) = props[0]
        if win32com.mapi.mapitags.PROP_TYPE(tag) == win32com.mapi.mapitags.PT_ERROR :
            raise TypeError('Error opening root folder of %s' % self.storename)
        return mapifolder (self, self.msgstore.OpenEntry (eid, None, win32com.mapi.mapi.MAPI_MODIFY))
    
    def OpenInbox (self) :
        cbEntryID, pEntryID = self.msgstore.GetReceiveFolder(None, 0, None)
        return mapifolder (self.msgstore.OpenEntry(cbEntryID, pEntryID, win32com.mapi.mapi.MAPI_MODIFY))        
        
    def OpenSpecialFolder (self, FolderID) :
        pInbox = self.OpenInbox()
        hr, props = pInbox.Folder().GetProps((FolderID), 0)
        (tag, eid) = props[0]
        if win32com.mapi.mapitags.PROP_TYPE(tag) == win32com.mapi.mapitags.PT_ERROR :
            raise TypeError('Error opening special folder of %s' % self.storename)
        return mapifolder (self, self.msgstore.OpenEntry (eid, None, win32com.mapi.mapi.MAPI_MODIFY))
        
    def OpenCalendar (self) :
        return self.OpenSpecialFolder (win32com.mapi.mapitags.PR_IPM_APPOINTMENT_ENTRYID)
        


        