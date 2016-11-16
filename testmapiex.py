import mapiex

# Change these to something valid before running !!!!
eml = 'PATH TO EML FILE'
storename = 'VALID MAPI STORE NAME'
dirname = 'VALID MAPI DIRECTORY NAME RO CREATE IN STORE'

eml = 'c:/Users/C07056/Documents/temp/1.eml'
storename = 'temp'
dirname = 'Crud'

# Test of enumeration of the message stores that are opened
def EnumerateMessageStore (MAPI) :
    messagestorestable = MAPI._GetContents ()
    while True:
        rows = messagestorestable.QueryRows(1, 0)
        #if this is the last row then stop
        if len(rows) != 1:
            break
        row = rows[0]
        # unpack the row and print name of the message store
        (eid_tag, eid), (name_tag, name), (def_store_tag, def_store) = row
        print("Store Name : %s" % name)

# Enumeration of the subfolders of a messagestore      
def EnumerateSubFolders (folder) :
    f = folder.GetFirstSubFolder ()
    while f != None :
        print("SubFolder : %s" % f.name)
        EnumerateSubFolders (f)
        f = folder.GetNextSubFolder ()
    

MAPI = mapiex.mapi()

print("Profile Name : %s " % MAPI.GetProfileName())
print ("Profile Email : %s" % MAPI.GetProfileEmail())

EnumerateMessageStore (MAPI)

MAPI.OpenMessageStore (storename)
rootfolder = MAPI.OpenRootFolder ()

EnumerateSubFolders (rootfolder)

folder = rootfolder.CreateSubFolder (dirname)
if folder == None :
    print("Can't open folder %s" % dirname)

folder.ImportEML (eml)

