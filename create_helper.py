# setup.py
from distutils.core import setup
import py2exe
import sys
import platform

def main () :
    
    if platform.architecture()[0] == "32bit":
        helper = 'helper32'
    else :
        helper = 'helper64'
        
    py2exe_options = dict(
        packages = [],
        optimize=2,
        compressed=True, # uncompressed may or may not have a faster startup
        bundle_files=3,
        dist_dir=helper,
        )
    
    setup(console=['eml2pst.py'],
          name = "EML2PST",
          version = "0.1",
          description = "A simple EML to PST conversion utility",
          options={"py2exe": py2exe_options},
        )

if len(sys.argv) == 1:
    sys.argv.append("py2exe")
    
main()