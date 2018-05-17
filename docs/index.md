Written by David Bateman <dbateman@free.fr> based on [nlconverter](https://github.com/kdeldycke/nlconverter)
by Hugues Bernard <hugues.bernard@gmail.com>. German translation by Faryan Rezagholi

NSF2X relies on the Windows COM interface to Lotus Notes for the conversion to EML and MBOX formats.
For the PST format it relies on the COM interface to Outlook and the Windows MAPI API. For this reason
NSF2X is a Windows only program and the user must have Lotus Notes installed, and optionally Outlook
for the conversion to the PST format.

NSF2X is licensed under the GPL v2 and the source code is available on
[https://github.com/adb014/nsf2x](https://github.com/adb014/nsf2x). The code is written in Python, but
compiled versions are available for users that do not want to install Python.

## Features of NSF2X
- Exports from NSF files to EML, MBOX and PST formats
- Exports the mail from Lotus Notes NSF files in MIME format keeping all layout and attachments
- Capable of reading the encrypted mails in the NSF files, removing the Lotus encryption and
  reencrypting with the users Exchange Certificate in RC2, 3DES, AES128 or AES256 formats
- Supports both Full and Click To Run (AKA Office 365) versions of Oulook
- Supports mixed 32 and 64 bit installations of Lotus Notes and Outlook
- Supports unicode filenames (ie. Accents in the NSF and PST filenames)
- Multi-lingual, though only English, French and German (partial) translations currently exist.

## Downloading NSF2X
The latest versions of the installers are available from
[https://github.com/adb014/nsf2x/releases/latest](https://github.com/adb014/nsf2x/releases/latest).

You should select the version based on whether your version of Lotus Notes is 32 or 64 bit. For 32bit
versions of Notes select the 'x86' version and for 64bit versions select the 'amd64' version.

## More Information
For people wanting to modify NSF2X, development notes of NSF2X are available in the
[README.dev](https://github.com/adb014/nsf2x/blob/master/README.dev) file