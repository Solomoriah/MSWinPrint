#!/usr/bin/env python
# Setup for MSWinPrint
# Modified from samples by Alex Martelli.

from distutils.core import setup

longdesc  =  """
MSWinPrint.py defines a class and utility functions for use by programs 
needing to print complex output on Windows 2K/XP hosts.

document is a class for creating and running print jobs.  Presently, the 
source is the only documentation for this class.

listprinters() returns a list of printer names.  The default printer is the 
first element of the list, and all other printers follow in alphabetical order.

desc(printer) returns a dictionary containing the descriptive fields
for the named printer.  

getfont(name, size) returns a win32ui font object for the named
font scaled to the given size.  Font substitution may have been
done by Windows, so don't be surprised if you don't get what you
asked for.

Development versions of this module may be found on **Github** at:

https://github.com/Solomoriah/MSWinPrint
"""

setup(
    name = "MSWinPrint",
    version = "1.1",
    description = "MSWinPrint",
    long_description = longdesc,
    author = "Chris Gonnerman",
    author_email = "chris@gonnerman.org",
    url = "http://newcenturycomputers.net/projects/mswinprint.html",
    py_modules = [ "MSWinPrint" ],
    keywords = "windows printing",

    classifiers = [
        "Programming Language :: Python",
        "Development Status :: 5 - Production/Stable",
        "Environment :: Other Environment",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: BSD License",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: Microsoft :: Windows :: Windows Server 2003",
        "Operating System :: Microsoft :: Windows :: Windows Server 2008",
        "Operating System :: Microsoft :: Windows :: Windows Vista",
        "Operating System :: Microsoft :: Windows :: Windows XP",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Printing",
    ],
)

# end of file.
