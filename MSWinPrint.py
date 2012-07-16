# MSWinPrint.py
# Copyright (c) 2004-2012 Chris Gonnerman
# All rights reserved.
# 
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
# 
# Redistributions of source code must retain the above copyright notice, this
# list of conditions and the following disclaimer.  Redistributions in binary
# form must reproduce the above copyright notice, this list of conditions and
# the following disclaimer in the documentation and/or other materials
# provided with the distribution.
# 
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

"""MSWinPrint -- Windows Printing Subsystem

MSWinPrint defines a class for use by programs needing to print complex
output on Windows 9x/2K/XP hosts.

document is a class for creating and running print jobs.

listprinters() returns a list of printer names.  The default printer is the 
first element of the list, and all other printers follow in alphabetical order.

desc(printer) returns a dictionary containing the descriptive fields
for the named printer.  

getfont(name, size) returns a win32ui font object for the named
font scaled to the given size.  Font substitution may have been
done by Windows, so don't be surprised if you don't get what you
asked for.
"""

# "constants" for use with printer setup calls

HORZRES = 8
VERTRES = 10
LOGPIXELSX = 88
LOGPIXELSY = 90
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111

import win32gui, win32ui, win32print, win32con

try:
    from PIL import ImageWin
except:
    ImageWin = None

scale_factor = 20

prdict = None

paper_sizes = {
    "letter":       1,
    "lettersmall":  2,
    "tabloid":      3,
    "ledger":       4,
    "legal":        5,
    "statement":    6,
    "executive":    7,
    "a3":           8,
    "a4":           9,
    "envelope9":   19,
    "envelope10":  20,
    "envelope11":  21,
    "envelope12":  22,
    "envelope14":  23,
    "fanfold":     39,
}

orientations = {
    "portrait":     1,
    "landscape":    2,
}

duplexes = {
    "normal":       1,
    "none":         1,
    "long":         2,
    "short":        3,
}

class document:

    def __init__(self, printer = None, papersize = None, orientation = None, duplex = None):
        self.dc = None
        self.font = None
        self.printer = printer
        self.papersize = papersize
        self.orientation = orientation
        self.page = 0
        self.duplex = duplex

    def scalepos(self, pos):
        rc = []
        for i in range(len(pos)):
            p = pos[i]
            if i % 2:
                p *= -1
            rc.append(int(p * scale_factor))
        return tuple(rc)

    def begin_document(self, desc = "MSWinPrint.py print job"):

        # open the printer
        if self.printer is None:
            self.printer = win32print.GetDefaultPrinter()
        self.hprinter = win32print.OpenPrinter(self.printer)

        # load default settings
        devmode = win32print.GetPrinter(self.hprinter, 8)["pDevMode"]

        # change paper size and orientation
        if self.papersize is not None:
            if type(self.papersize) is int:
                devmode.PaperSize = self.papersize
            else:
                devmode.PaperSize = paper_sizes[self.papersize]
        if self.orientation is not None:
            devmode.Orientation = orientations[self.orientation]
        if self.duplex is not None:
            devmode.Duplex = duplexes[self.duplex]

        # create dc using new settings
        self.hdc = win32gui.CreateDC("WINSPOOL", self.printer, devmode)
        self.dc = win32ui.CreateDCFromHandle(self.hdc)

        # self.dc = win32ui.CreateDC()
        # if self.printer is not None:
        #     self.dc.CreatePrinterDC(self.printer)
        # else:
        #     self.dc.CreatePrinterDC()

        self.dc.SetMapMode(win32con.MM_TWIPS) # hundredths of inches
        self.dc.StartDoc(desc)
        self.pen = win32ui.CreatePen(0, int(scale_factor), 0L)
        self.dc.SelectObject(self.pen)
        self.page = 1

    def end_document(self):
        if self.page == 0:
            return # document was never started
        self.dc.EndDoc()
        del self.dc

    def end_page(self):
        if self.page == 0:
            return # nothing on the page
        self.dc.EndPage()
        self.page += 1

    def getsize(self):
        if self.page == 0:
            self.begin_document()
        # returns printable (width, height) in points
        width = float(self.dc.GetDeviceCaps(HORZRES)) * (72.0 / self.dc.GetDeviceCaps(LOGPIXELSX))
        height = float(self.dc.GetDeviceCaps(VERTRES)) * (72.0 / self.dc.GetDeviceCaps(LOGPIXELSY))
        return width, height

    def line(self, from_, to):
        if self.page == 0:
            self.begin_document()
        self.dc.MoveTo(self.scalepos(from_))
        self.dc.LineTo(self.scalepos(to))

    def rectangle(self, box):
        if self.page == 0:
            self.begin_document()
        self.dc.MoveTo(self.scalepos((box[0], box[1])))
        self.dc.LineTo(self.scalepos((box[2], box[1])))
        self.dc.LineTo(self.scalepos((box[2], box[3])))
        self.dc.LineTo(self.scalepos((box[0], box[3])))
        self.dc.LineTo(self.scalepos((box[0], box[1])))

    def text(self, position, text):
        if self.page == 0:
            self.begin_document()
        self.dc.TextOut(scale_factor * position[0],
            -1 * scale_factor * position[1], text)

    def setfont(self, name, size, bold = None):
        if self.page == 0:
            self.begin_document()
        wt = 400
        if bold:
            wt = 700
        self.font = getfont(name, size, wt)
        self.dc.SelectObject(self.font)

    def image(self, position, image, size):
        "print PIL image at position with given size"
        if ImageWin is None:
            raise NotImplementedError, "PIL required for image method"
        if self.page == 0:
            self.begin_document()
        dib = ImageWin.Dib(image)
        endpos = (position[0] + size[0], position[1] + size[1])
        dest = (position[0] * scale_factor, 
               -1 * position[1] * scale_factor,
               endpos[0] * scale_factor, 
               -1 * endpos[1] * scale_factor)
        dib.draw(self.hdc, dest)

    def setink(self, ink):
        pass

    def setfill(self, onoff):
        pass

def build_dict():
    global prdict
    lst = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_CONNECTIONS
        + win32print.PRINTER_ENUM_LOCAL)
    prdict = {}
    for flags, description, name, comment in lst:
        prdict[name] = {}
        prdict[name]["flags"] = flags
        prdict[name]["description"] = description
        prdict[name]["comment"] = comment

def listprinters():
    dft = win32print.GetDefaultPrinter()
    if prdict is None:
        build_dict()
    keys = prdict.keys()
    keys.sort()
    rc = [ dft ]
    for k in keys:
        if k != dft:
            rc.append(k)
    return rc

def desc(name):
    if prdict == None:
        listprinters()
    return prdict[name]

def getfont(name, size, weight = 400):
    return win32ui.CreateFont({
        "name": name,
        "height": scale_factor * size,
        "weight": weight,
    })


if __name__ == "__main__":

    doc = document(orientation = "landscape")
    doc.begin_document()
    doc.setfont("Arial", 32)
    doc.text((72, 72), "Testing...")
    doc.text((72, 72+48), "Testing #2")
    doc.rectangle((72, 72, 72*6, 72*3))
    doc.line((72, 72), (72*6, 72*3))
    doc.end_document()

# end of file.
