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

# Copyright (C) 2009 Free Software Foundation
# Author : David Bateman <dbateman@free.fr>

# This script is used to create a standalone binary package of NSF2X
# with python3

import distutils.core
import py2exe
import os
import zipfile
import sys
import subprocess
import platform

class Target(object):
    '''Target is the baseclass for all executables that are created.
    It defines properties that are shared by all of them.
    '''
    def __init__(self, **kw):
        self.__dict__.update(kw)

        # the VersionInfo resource, uncomment and fill in those items
        # that make sense:
        
        # The 'version' attribute MUST be defined, otherwise no versioninfo will be built:
        # self.version = "1.0"
        
        # self.company_name = "Company Name"
        # self.copyright = "Copyright Company Name © 2013"
        # self.legal_copyright = "Copyright Company Name © 2013"
        # self.legal_trademark = ""
        # self.product_version = "1.0.0.0"
        # self.product_name = "Product Name"

        # self.private_build = "foo"
        # self.special_build = "bar"

    def copy(self):
        return Target(**self.__dict__)

    def __setitem__(self, name, value):
        self.__dict__[name] = value

def which(program):
    def is_exe(fpath):
        return os.path.isfile(fpath) and os.access(fpath, os.X_OK)

    fpath, fname = os.path.split(program)
    if fpath:
        if is_exe(program):
            return program
    else:
        for path in os.environ["PATH"].split(os.pathsep):
            path = path.strip('"')
            exe_file = os.path.join(path, program)
            if is_exe(exe_file):
                return exe_file

    return None
 
def find_all_files_in_dir(directory):
    ret = []
    for root, dir, files in os.walk(directory) :
        _files = ()
        for file in files :
            _files += (os.path.join(root, file),)
        if _files :
            ret += [(root, _files)]
    return ret
    
def main () :
    if platform.architecture()[0] == "32bit":
        bitness = 'x86'
    else :
        bitness = 'amd64'
        
    RT_BITMAP = 2
    RT_MANIFEST = 24

    version = "1.3.1"
    author = "dbateman@free.fr"
    description="NSF2X - Converts Lotus NSF files to EML, MBOX or PST files..."
    
    # A manifest which specifies the executionlevel
    # and windows common-controls library version 6

    manifest_template = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
      <assemblyIdentity
        version="5.0.0.0"
        processorArchitecture="*"
        name="%(prog)s"
        type="win32"
      />
      <description>%(prog)s</description>
      <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
        <security>
          <requestedPrivileges>
            <requestedExecutionLevel
                level="%(level)s"
                uiAccess="false">
            </requestedExecutionLevel>
          </requestedPrivileges>
        </security>
      </trustInfo>
      <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.Common-Controls"
                version="6.0.0.0"
                processorArchitecture="*"
                publicKeyToken="6595b64144ccf1df"
                language="*"
            />
        </dependentAssembly>
      </dependency>
    </assembly>
    '''

    nsf2x = Target(
        # We can extend or override the VersionInfo of the base class:
        version = version,
        file_description = description,
        # comments = "Some Comments",
        internal_name = "NSF2X",

        script="nsf2x.py", # path of the main script

        # Allows to specify the basename of the executable, if different from 'nsf2x'
        # dest_base = "nsf2x",

        # Icon resources:[(resource_id, path to .ico file), ...]
        # icon_resources=[(1, r"nsf2x.ico")]

        other_resources = [(RT_MANIFEST, 1, (manifest_template % dict(prog="nsf2x", level="asInvoker")).encode("utf-8")),
        # for bitmap resources, the first 14 bytes must be skipped when reading the file:
        #                    (RT_BITMAP, 1, open("bitmap.bmp", "rb").read()[14:]),
                          ]
        )


    # ``zipfile`` and ``bundle_files`` options explained:
    # ===================================================
    #
    # zipfile is the Python runtime library for your exe/dll-files; it
    # contains in a ziparchive the modules needed as compiled bytecode.
    #
    # If 'zipfile=None' is used, the runtime library is appended to the
    # exe/dll-files (which will then grow quite large), otherwise the
    # zipfile option should be set to a pathname relative to the exe/dll
    # files, and a library-file shared by all executables will be created.
    #
    # The py2exe runtime *can* use extension module by directly importing
    # the from a zip-archive - without the need to unpack them to the file
    # system.  The bundle_files option specifies where the extension modules,
    # the python dll itself, and other needed dlls are put.
    #
    # bundle_files == 3:
    #     Extension modules, the Python dll and other needed dlls are
    #     copied into the directory where the zipfile or the exe/dll files
    #     are created, and loaded in the normal way.
    #
    # bundle_files == 2:
    #     Extension modules are put into the library ziparchive and loaded
    #     from it directly.
    #     The Python dll and any other needed dlls are copied into the
    #     directory where the zipfile or the exe/dll files are created,
    #     and loaded in the normal way.
    #
    # bundle_files == 1:
    #     Extension modules and the Python dll are put into the zipfile or
    #     the exe/dll files, and everything is loaded without unpacking to
    #     the file system.  This does not work for some dlls, so use with
    #     caution.
    #
    # bundle_files == 0:
    #     Extension modules, the Python dll, and other needed dlls are put
    #     into the zipfile or the exe/dll files, and everything is loaded
    #     without unpacking to the file system.  This does not work for
    #     some dlls, so use with caution.


    py2exe_options = dict(
        packages = [],
    ##    excludes = "tof_specials Tkinter".split(),
    ##    ignores = "dotblas gnosis.xml.pickle.parsers._cexpat mx.DateTime".split(),
    ##    dll_excludes = "MSVCP90.dll mswsock.dll powrprof.dll".split(),
        optimize=0,
        compressed=False, # uncompressed may or may not have a faster startup
        bundle_files=3,
        dist_dir='dist',
        )


    # Some options can be overridden by command line options...
    distutils.core.setup(name="NSF2X",

          # windows subsystem executables (no console)
          windows=[nsf2x],
          
          # version of the program
          version=version,
          
          # short description of the program
          description=description,
          
          # author contact details
          author=author,
          url='mailto:' + author,
          
          # data files to include
          data_files=[(".", ("README.txt", "LICENSE")), 
                      ("src", ("create_exe.py", "create_helper.py", "eml2pst.py",
                               "nsf2x.py", "mapiex.py", "testmapiex.py",
                               "nsf2x.nsi", "README.dev"))] +
                        find_all_files_in_dir('locale') +
                        find_all_files_in_dir('helper32') +
                        find_all_files_in_dir('helper64'),

          # py2exe options
          zipfile=None,
          options={"py2exe": py2exe_options},
          )
    
    zf = zipfile.ZipFile("nsf2x-" + version + "-" + bitness + ".zip", "w", zipfile.ZIP_DEFLATED)
    for root, dir, files in os.walk("dist") :
        for file in files :
            _file = os.path.join(root, file)
            _arcname = "nsf2x" + _file[4:]
            zf.write (_file, _arcname)
    zf.close()

    # Run NSIS to create the installer. Prefer the portable version if installed
    makensis=which("NSISPortable.exe")
    if not makensis :
        makensis=which("makensis.exe")
    if not makensis :
        raise "Can not find NSIS executable on your path"
    subprocess.call([makensis, "-DVERSION=" + version, "-DPUBLISHER=" + author,
                     "-DBITNESS=" + bitness, "nsf2x.nsi"])
    
if len(sys.argv) == 1:
    sys.argv.append("py2exe")
    
main()