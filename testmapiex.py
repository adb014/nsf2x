import mapiex

# Change these to something valid before running !!!!
eml = 'PATH TO EML FILE'
storename = 'VALID MAPI STORE NAME'
dirname = 'VALID MAPI DIRECTORY NAME TO CREATE IN STORE'

# Test of enumeration of the message stores that are opened
def EnumerateMessageStore (MAPI) :
    for name in MAPI.GetMessageStoreNames() :
        print("Store Name : %s" % name)

# Enumeration of the subfolders of a messagestore      
def EnumerateSubFolders (folder, indent="") :
    f = folder.GetFirstSubFolder ()
    while f != None :
        print("%sSubFolder : %s" % (indent, f.name))
        EnumerateSubFolders (f, indent + "  ")
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

