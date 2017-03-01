NSF2X : A Lotus Notes NSF to EML, MBOX and Outlook PST converter
dbateman@free.fr

Based on nlconverter (https://code.google.com/p/nlconverter/) by
Hugues Bernard <hugues.bernard@gmail.com>

NSF2X supports multiple languages, using GNU GetText based on your regional
settings. Only French and English are currently available, though as I'm not
a native French speaker, corrections are always welcome.

Quick Start
-----------
   0. Install/Unzip the downloaded package on a Windows machine
   1. Make a copy of the *.nsf files you want to convert to a temporary location
   2. Launch Lotus Notes
   3. Optionally, but recommended, launch Outlook
   4. Launch "nsf2x.exe"
   5. Enter the Lotus Notes password
   6. Press the "Open Session" to open the connection to Notes
   7. Select the output type: EML, MBOX or PST
   8. Modify the conversion options as wanted
   9. Enter the source path of the temporary location with the "*.nsf" files
  10. Enter the destination path to contain the converted files
  11. Press the "Convert" button to launch the conversion
  12. Progress is displayed in the title bar, errors in the window
  13. Enjoy
  14. Treat any exceptions manually

WARNING
-------
NSF2X can read all the encrypted mails that your Notes ID gives you access to. It
decrypts these mail in its output EML, and MBOX files, as NSF2X can't use the
Notes encryption in the EML files. So if you care about the security of your mail
archives, store them on encrypted disks after conversion. You've been warned.
However, a mail that is encrypted in Notes can be re-encrypted with the users 
certificate in the PST file, if you have their certificates.

NSF2X is also relatively slow for conversion to PST files, 10000 mails took me about
30 minutes to convert on a reasonable laptop. So those 2GB NSF files of yours
are going to take some time to convert to PST files !!! Writing to EML or MBOX is
about 3 to 5 times faster 
  
Details
-------

   0. Install/Unzip the downloaded package on a Windows machine
  -------------------------------------------------------------
   Well if you're reading this then you've probably already done this. The
   binary releases are available from
   
   https://github.com/adb014/nsf2x/releases
   
   If you haven't already guessed, this program relies on a number of Windows
   features, notably the COM interface of Notes and Outlook as well as Outlook
   itself. These being Windows only, no this program won't work under Linux.
   
   The installer permits NSF2X to be installed for the current users or all
   users. However, you must have administration privileges to be able to 
   install for all users. A single user install, doesn't need elevated 
   privileges.
   
   Please note that if you have Outlook 2013 or 2016 installed in "Click to Run"
   mode (aka "Office 365"), then NSF2X can not convert to PST files. You have 
   several choices in this case
   
   A/ Upgrade your installation of Office to the "Open license full download" 
   version. In this case the installation of Outlook will not be in "Click To
   Run" mode
   
   B/ Allow the NSF2X installer to patch the registry. In this case NSF2X must
   be installed with administrator privileges. This will allow NSF2X to run
   correctly, however if you have multiple versions of Outlook installed (For
   example Outlook 2007 and Outlook 2016) then the non "Click To Run" version 
   (in this case Outlook 2007) might fail in unexpected ways. 
   
   In this case after having used NSF2X to convert your archives, you should 
   uninstall it, letting the NSF2X uninstaller remove the modifications to the
   registry. See
   
   https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/
   
   for more information.
   
   C/ Continue the installation of NSF2X knowing that the conversion to PST 
   files will not be possible
   
   D/ Don't use NSF2X
   
   1. Make a copy of the *.nsf files you want to convert to a temporary location
  ------------------------------------------------------------------------------
   I make no guarantee that this program won't destroy your Lotus NSF archives,
   every file on your computer or kill your cat. That being said, I've used it
   for my needs without problems and so you too should be able to do so.

   The process of converting to EML files will modify the NSF files. The reason
   is that each message is converted to MIME within the NSF file even if it
   wasn't initially coded in MIME within Lotus Notes. For this reason it is
   better to make a copy of your NSF files and let NSF2X work on these copies.
   By doing this you'll minimize the risk of loss of data
   
   This needs to be done before Lotus Notes is running to ensure that Lotus
   doesn't prevent you from making a copy, as the archive is open in Notes.

   2. Launch Lotus Notes
  ----------------------
   NSF2X relies on Lotus Notes to do the heavy lifting for the conversion to
   EML files. NSF2X will try to use Lotus Notes even if you haven't launched it
   so this step is optional, but recommended

   3. Launch Outlook
  ------------------
   For the conversion to PST, Outlook is necessary. It isn't necessary for the
   conversion to EML or MBOX. NSF2X will start Outlook the first time it needs
   it, but if you have already started Outlook, NSF2X will load the EML files
   for conversion to PST faster.

   4. Launch "nsf2x.exe"
  ---------------------
   The file "nsf2x.exe" is the compiled version of the source code file
   "nsf2x.py". The code is written in Python 2.6 compatible code and compiled
   with the py2exe code so that you don't need python installed to run it.
   Though you'll need Python if you alter the "nsf2x.py" code
   
   5. Enter the Lotus Notes password
  ----------------------------------
   So that NSF2X can have access to Lotus Notes it needs your Lotus Notes
   password. Enter into the password box of NSF2X

   6. Press the "Open Session" to open the connection to Notes
  ------------------------------------------------------------
   Pressing this button opens the connection to Lotus Notes via a Windows 
   COM interface. At this point the password selection is  deactivated, but
   the source and destination path entries are activated.
   
   7. Select the output type: EML, MBOX or PST
  --------------------------------------------
   NSF2X can convert to EML, MBOX or PST formats. For each NSF file found in
   the source directory NSF2X does the following steps

   EML :
   .....
   For each NSF a sub-directory "<DestPath>/<NSFFileBasename>" is created. 
   where <NSFFileBasename> is the NSF file with the "*.nsf" termination
   removed. Under these sub-directories, the folder hierarchy of the NSF file 
   is recreated and each message of each folder is created in a separate 
   EML file.

   MBOX :
   ......
   There are two possibles means of treating the conversion of MBOX files
   In the first case, for each NSF file an MBOX file is created in <DestPath>, 
   with the ".nsf" termination replaced with ".mbox". Unfortunately the folder
   hierarchy is thrown away in the created MBOX file, and a flat hierarchy is 
   used. The advantage is only a single MBOX file is created for each NSF file.
   
   In the second case a folder hierarchy is created with an MBOX representing
   each NSF sub-directory, thus retaining the folder hierarchy. The downside
   is a large number of MBOX files is potentially created.

   PST :
   .....
   For each NSF file a PST file is created in <DestPath>, with the ".nsf"
   termination replaced with ".pst". The folder hierarchy in the NSF file is
   recreated in the PST file. Each message from the NSF file is saved to a
   temporary file in <DestPath> and then opened by Outlook and moved to the
   correct folder. At the end of the process the PST file is left open within
   Outlook. You can either close these PST files before moving them to their
   final location and reopen them in Outlook or create them directly in their
   final location. If you run NSF2X twice with the same source and destination, 
   the messages in the NSF file will be copied to the PST files twice.

   8. Modify the conversion options as wanted
  -------------------------------------------   
   Using the "Options" button the user can modify three parameters of NSF2X.
   The options that are be modified are discussed below
   
   Use different MBOXes for each sub-folder :
   ..........................................
   This option only concerns the conversion to MBOX format. The possible 
   options are
   
   No : A single MBOX file will be created and the sub-directory hierarchy 
   will be discarded

   YES : The sub-directory hierarchy is created using Windows folders and
   a separate MBOX file will be created for each sub-directory
   
   Treatment of encrypted Notes messages :
   .......................................
   This option concerns all conversion types. The possible options are
   
   None : The encryption status of all Notes mails is ignored and all mail is 
   saved without encryption
   
   RC2 40bit : Encrypt with the algorithm RC2-CBC with 40 bit. This algorithm
   is quite weak, but very portable. If you can avoid it you shouldn't use it
   
   3DES 168bit : Encrypt with DES EDE3 CBC or Triple DES with 168 bits. Some
   people consider this encryption as being fragile, but it offers a compromise
   of security and compatibility 
   
   AES 128 bit : Encrypt with a modern AES 128 bit encryption
   
   AES 256 bit : Encrypt with a modern AES 256 bit encryption 
   
   As the Windows certificate store is used with the users default
   certificate for exportation to mail clients that don't use the same Windows
   certificate store (for example Thunderbird), the user might need to export
   their certificate and re-import it into their mail client to be able to
   read the encrypted mails. This should not be the case for Outlook on the
   same machine.

   Error logging level
   ...................
   This option concerns all conversion types. The possible options are

   Error : As well as the messages for the progress of NSF2X, display error
   messages
   
   Warning : As well as the normal and error messages, display warnings for
   NSF2X
   
   Information : Display all messages, can be rather verbose
   
   Number of exceptions before giving up
   .....................................
   This option concerns all conversion types. The possible options are 
   
   1 : A single exception will cause NSF2X to stop. This is only useful while 
   debugging
   
   10 : Ten exceptions are allowed before NSF2X stops
   
   100 : One hundred exceptions are allowed beofre NSF2X stops
   
   Infinite : NSF2X will run until the end of the conversion regardless of the
   number of exceptions.
   
   9. Enter the source path of the temporary location with the "*.nsf" files
  --------------------------------------------------------------------------
   Clicking on the source directory entry will open a dialog to select a
   source directory, with the NSF files.

  10. Enter the destination path to contain the converted files
  -------------------------------------------------------------
   Clicking on the destination directory will open a dialog to select a
   destination directory
   
  11. Press the "Convert" button to launch the conversion
  -------------------------------------------------------
   The conversion process is launch and the UI is disabled, leaving only the
   option to "Stop" the process
   
  12. Progress is displayed in the title bar, errors in the window
  ----------------------------------------------------------------
   Error messages, information and warnings are printed in the window of
   NSF2X. These will be useful to debug any problems you have.
   
   If you get the error "File does not exist error (259)" repeatedly from the
   MIMEConvertCDParts function while converting to MIME, then a workaround 
   is to convert NSF file by NSF file, restarting NSF2X in between. This 
   seems to be due to a weird inter-action between Lotus and Outlook that 
   is unresolved (and not in my code).
   
  13. Enjoy
  ---------
   For large NSF file you're better off going and getting a coffee or going
   to bed. You can however lock the screen without interrupting the conversion
   as long as the NSF and PST files are in locations that NSF2X can access
   with the screen locked
   
   14. Treat any exceptions manually
   ---------------------------------
   Some mails in Lotus notes might be malformed and not capable of being
   transformed into MIME by Lotus. It is also possible that an unexpected
   error might occur. Hopefully, this will be very rare or not occur at all,
   but if it does the output of NSF2X will tell the user the number of 
   exceptions of this type at the end of its execution of each NSF file. The
   number of "untreated" documents are messages that NSF2X knows are not
   mails and are not treated (for example delivery failure reports, replies
   to "Appointments", etc)
   
   The messages that weren't transferred from Lotus will have their subjects 
   printed to the log window with the prefix
   
   #### Subject :  ...
   
   This can be used to identify the untransferred messages in the Lotus NSF 
   interface. Don't despair these messages are not lost and can be transferred
   manually.
   
   To manually transfer these files, the message can be dragged to the Windows
   Desktop. Lotus will then save them to an EML file via a different mechanism
   used by NSF2X. These EML files can be easily treated for the EML or MBOX
   message formats of NSF2X. 
   
   For the PST format the manner to import these messages manually into the PST
   file is to 
   
   1. Launch CMD.EXE to get a Windows commandline
   2. Identify the location of the exported EML file. For example

      C:\Users\Me\Desktop\Message.eml
      
   3. Ensure that the executable OUTLOOK.EXE is on your path
   4. Type the command
   
      OUTLOOK.EXE /eml C:\Users\Me\Desktop\Message.eml
      
   This will pop up a window in Outlook. At this point the message is not in any
   Outlook OST or PST file. However the message can be dragged to the desired folder
   within Outlook.

Copyright
---------

As a derivative program it inherits the same license (i.e. GPL v2)

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
# Author : Hugues Bernard <hugues.bernard@gmail.com>
# Author : David Bateman <dbateman@free.fr>
