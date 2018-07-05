# NSF2X
A Lotus Notes NSF to EML, MBOX and Outlook PST converter by [David Bateman](dbateman@free.fr) based on [nlconverter](https://code.google.com/p/nlconverter/) by [Hugues Bernard](hugues.bernard@gmail.com). The code is written in Python, but compiled versions are available for users that do not want to install Python.
---
NSF2X relies on the Windows COM interface to Lotus Notes for the conversion to EML and MBOX formats. For the PST format it relies on the COM interface to Outlook and the Windows MAPI API. For this reason NSF2X is a Windows only program and the user must have Lotus Notes installed, and optionally Outlook for the conversion to the PST format.

## Features of NSF2X
  * Exports from NSF files to EML, MBOX and PST formats
  * Exports the mail from Lotus Notes NSF files in MIME format keeping all layout and attachments
  * Capable of reading the encrypted mails in the NSF files, removing the Lotus encryption and reencrypting with the users Exchange Certificate in RC2, 3DES, AES128 or AES256 formats
  * Supports both Full and Click To Run (AKA Office 365) versions of Oulook
  * Supports mixed 32 and 64 bit installations of Lotus Notes and Outlook
  * Supports unicode filenames (ie. Accents in the NSF and PST filenames)
  * Multi-lingual, though only English, French and German translations currently exist

You should select the version based on whether your version of Lotus Notes is 32 or 64 bit. For 32bit versions of Notes select the 'x86' version and for 64bit versions select the 'amd64' version.

Download the latest installers from the [Releases](https://github.com/adb014/nsf2x/releases/latest) section of the site!
For people wanting to modify NSF2X, development notes of NSF2X are available in the [README.dev](./README.dev) file.

## Quick Start
  * Install/Unzip the downloaded package on a Windows machine
  * Make a copy of the nsf-files you want to convert to a temporary location
  * Launch Lotus Notes
  * Optionally, but recommended, launch Outlook
  * Launch "nsf2x.exe"
  * Enter the Lotus Notes password
  * Press the "Open Session" button to open the connection to Notes
  * Select the output type: EML, MBOX or PST
  * Modify the conversion options as wanted
  * Enter the source path of the temporary location with the nsf-files
  * Enter the destination path to contain the converted files
  * Press the "Convert" button to launch the conversion
  * Progress is displayed in the title bar, errors in the window
  * Enjoy
  * Treat any exceptions manually

## Warning
NSF2X can read all the encrypted mails that your Notes ID gives you access to. It can decrypt these mail in its output EML, and MBOX files, as NSF2X can't use the Notes encryption in the EML files. So if you care about the security of your mail
archives, store them on encrypted disks after conversion. You've been warned. However, a mail that is encrypted in Notes can be re-encrypted with the users certificate, if you have their certificates installed in the Microsoft Crypto Store. These encrypted mails are directly useable by Outlook but you might have to export the users certificate to read the EML or MBOX exported files.

NSF2X is also relatively slow for conversion to PST files, 10000 mails took me about 30 minutes to convert on a reasonable laptop. So those 2GB NSF files of yours are going to take some time to convert to PST files !!! Writing to EML or MBOX is about 3 to 5 times faster

## Details
### Install/Unzip the downloaded package on a Windows machine
Well if you're reading this then you've probably already done this. The binary releases are available from [Releases](./releases/latest)

If you haven't already guessed, this program relies on a number of Windows features, notably the COM interface of Notes and Outlook as well as Outlook itself. These being Windows only, no this program won't work under Linux.

The installer permits NSF2X to be installed for the current users or all users. However, you must have administration privileges to be able to install for all users. A single user install, doesn't need elevated privileges.

NSF2X is supplied in both 32bit and 64bit versions. The version used should match the bitness of the version of Lotus Notes that is used. In the case of conversion to an Outlook PST file it is possible to have a version of Outlook with a different bitness to Lotus Notes. In that case NSF2X will export the mail to a set of temporary EML files and then call an external helper program of the right bitness for Outlook to allow the conversion.

### Make a copy of the nsf-files you want to convert to a temporary location
I make no guarantee that this program won't destroy your Lotus NSF archives, every file on your computer or kill your cat. That being said, I've used it for my needs without problems and so you too should be able to do so.

The process of converting to EML files will modify the NSF files. The reason is that each message is converted to MIME within the NSF file even if it wasn't initially coded in MIME within Lotus Notes. For this reason it is better to make a copy of your NSF files and let NSF2X work on these copies. By doing this you'll minimize the risk of loss of data

This needs to be done before Lotus Notes is running to ensure that Lotus doesn't prevent you from making a copy, as the archive is open in Notes.

### Launch Lotus Notes
NSF2X relies on Lotus Notes to do the heavy lifting for the conversion to EML files. NSF2X will try to use Lotus Notes even if you haven't launched it so this step is optional, but recommended

### Launch Outlook
For the conversion to PST, Outlook is necessary. It isn't necessary for the conversion to EML or MBOX. NSF2X will start Outlook the first time it needs it, but if you have already started Outlook, NSF2X will load the EML files for conversion to PST faster.

### Launch "nsf2x.exe"
The file "nsf2x.exe" is the compiled version of the source code file "nsf2x.py". The code is written in Python 2.6 compatible code and compiled with the py2exe code so that you don't need python installed to run it. Though you'll need Python if you alter the "nsf2x.py" code

### Enter the Lotus Notes password
So that NSF2X can have access to Lotus Notes it needs your Lotus Notes password. Enter into the password box of NSF2X

### Press the "Open Session" button to open the connection to Notes
Pressing this button opens the connection to Lotus Notes via a Windows COM interface. At this point the password selection is deactivated, but the source and destination path entries are activated.

### Select the output type: EML, MBOX or PST
NSF2X can convert to EML, MBOX or PST formats. For each NSF file found in the source directory NSF2X does the following steps
#### EML
For each NSF a sub-directory "<DestPath>/<NSFFileBasename>" is created. where <NSFFileBasename> is the NSF file with the `.nsf` termination
removed. Under these sub-directories, the folder hierarchy of the NSF file is recreated and each message of each folder is created in a separate EML file.
#### MBOX
There are two possibles means of treating the conversion of MBOX files In the first case, for each NSF file an MBOX file is created in <DestPath>, with the ".nsf" termination replaced with ".mbox". Unfortunately the folder hierarchy is thrown away in the created MBOX file, and a flat hierarchy is used. The advantage is only a single MBOX file is created for each NSF file.
In the second case a folder hierarchy is created with an MBOX representing each NSF sub-directory, thus retaining the folder hierarchy. The downside is a large number of MBOX files is potentially created.
#### PST
For each NSF file a PST file is created in <DestPath>, with the ".nsf" termination replaced with ".pst". The folder hierarchy in the NSF file is recreated in the PST file. Each message from the NSF file is saved to a temporary file in <DestPath> and then opened by Outlook and moved to the correct folder. At the end of the process the PST file is left open within Outlook. You can either close these PST files before moving them to their final location and reopen them in Outlook or create them directly in their final location. If you run NSF2X twice with the same source and destination, the messages in the NSF file will be copied to the PST files twice.

### Modify the conversion options as wanted
Using the "Options" button the user can modify five parameters of NSF2X. The options that are be modified are discussed below

### Use different MBOXes for each sub-folder
This option only concerns the conversion to MBOX format. The possible options are
  * No : A single MBOX file will be created and the sub-directory hierarchy will be discarded
  * YES : The sub-directory hierarchy is created using Windows folders and a separate MBOX file will be created for each sub-directory

### Treatment of encrypted Notes messages
This option concerns all conversion types. The possible options are
  * None : The encryption status of all Notes mails is ignored and all mail is saved without encryption. This is useful if you can't export the user certificate from the Microsoft cryptographic store and therefore can't read encrypted EML and MBOX mails with your mail client.
  * RC2 40bit : Encrypt with the algorithm RC2-CBC with 40 bit. This algorithm is quite weak, but very portable. If you can avoid it you shouldn't use it
  * 3DES 168bit : Encrypt with DES EDE3 CBC or Triple DES with 168 bits. Some people consider this encryption as being fragile, but it offers a compromise of security and compatibility
  * AES 128 bit : Encrypt with a modern AES 128 bit encryption
  * AES 256 bit : Encrypt with a modern AES 256 bit encryption

As the Windows certificate store is used with the users default certificate for exportation to mail clients that don't use the same Windows certificate store (for example Thunderbird), the user might need to export their certificate and re-import it into their mail client to be able to read the encrypted mails. This should not be the case for Outlook on the same machine. The certificates can be exported via Outlook or Internet Explorer in a variety of formats.

### Error logging level
This option concerns all conversion types. The possible options are
  * Error : As well as the messages for the progress of NSF2X, display error
   messages
  * Warning : As well as the normal and error messages, display warnings for
   NSF2X
  * Information : Display all messages, can be rather verbose

### Number of exceptions before giving up
This option concerns all conversion types. The possible options are
   * 1 : A single exception will cause NSF2X to stop. This is only useful while debugging
   * 10 : Ten exceptions are allowed before NSF2X stops
   * 100 : One hundred exceptions are allowed beofre NSF2X stops
   * Infinite : NSF2X will run until the end of the conversion regardless of the number of exceptions.

### Always use external PST helper function
This options concerns conversion to PST format. The possible options are
  * Yes : The mail will all be stored to a temporary location and an external helper function will be called for the conversion of these EML files to the PST format.
If you repeatly get the message "File does not exist error (259)" repeatedly from the MIMEConvertCDParts function, then this option can be used to avoid it. The downside of this option is that additional disk space if needed to store the temporary EML files.

  * No : If Outlook is the same bitness as NSF2X, then NSF2X will convert directly to the PST format. Otherwise an external helper function will be used.

### Enter the source path of the temporary location with the nsf-files
Clicking on the source directory entry will open a dialog to select a source directory, with the NSF files.

### Enter the destination path to contain the converted files
Clicking on the destination directory will open a dialog to select a destination directory

### Press the "Convert" button to launch the conversion
The conversion process is launch and the UI is disabled, leaving only the option to "Stop" the process

### Progress is displayed in the title bar, errors in the window
Error messages, information and warnings are printed in the window of NSF2X. These will be useful to debug any problems you have.

If you get the error "File does not exist error (259)" repeatedly from the MIMEConvertCDParts function while converting to MIME, then a workaround is to convert NSF file by NSF file, restarting NSF2X in between. This seems to be due to a weird inter-action between Lotus and Outlook that is unresolved (and not in my code).

### Enjoy
For large NSF file you're better off going and getting a coffee or going to bed. You can however lock the screen without interrupting the conversion as long as the NSF and PST files are in locations that NSF2X can access with the screen locked

### Treat any exceptions manually
Some mails in Lotus notes might be malformed and not capable of being transformed into MIME by Lotus. It is also possible that an unexpected error might occur. Hopefully, this will be very rare or not occur at all, but if it does the output of NSF2X will tell the user the number of exceptions of this type at the end of its execution of each NSF file. The number of "untreated" documents are messages that NSF2X knows are not mails and are not treated (for example delivery failure reports, replies to "Appointments", etc)

The messages that weren't transferred from Lotus will have their subjects printed to the log window with the prefix `Subject :  ...`

This can be used to identify the untransferred messages in the Lotus NSF interface. Don't despair these messages are not lost and can be transferred manually.

To manually transfer these files, the message can be dragged to the Windows Desktop. Lotus will then save them to an EML file via a different mechanism than used by NSF2X. These EML files can be easily treated for the EML or MBOX message formats of NSF2X.

For the PST format the manner to import these messages manually into the PST file is to
  * Launch CMD.EXE to get a Windows commandline
  * Identify the location of the exported EML file. For example
  `C:\Users\Me\Desktop\Message.eml`
  * Ensure that the executable OUTLOOK.EXE is on your path
  * Type the command
  `OUTLOOK.EXE /eml C:\Users\Me\Desktop\Message.eml`

This will pop up a window in Outlook. At this point the message is not in any Outlook OST or PST file. However the message can be dragged to the desired folder within Outlook.
