
#Tulip Cell

##Introduction

Tulip Cell is an Excel add-in that provides the technical analysis functions
from the [Tulip Indicators](https://tulipindicators.org) library.


##Building

Building is a pain. I suggest you grab the installer from [the
website](https://tulipcell.org).

If you really want to build it, you'll need the Tulip Indicators code first. I
suggest getting the MinGW compiler. You'll want both the 64-bit and 32-bit
versions. You can look at the `Makefile`, `build32.bat`, and `build64.bat` to
get an idea of how I do things. You'll also want Tcl installed to generate the
Excel VBA wrapper from `create_vba.tcl`.

Building goes like this:

0. Set your Tulip Indicators path in the `Makefile`

0. Run `build32.bat` from a 32-bit compiler environment. This will make
   `tulipcell32.dll`.

0. Run `build64.bat` from a 64-bit compiler environment. This will make
   `tulipcell64.dll`.

0. Run `create_vba.tcl`. This writes out `vba.txt`. This is the VBA code that
   you'll need to put in an Excel add-in.

0. Put the code in an Excel add-in. Save as both `.xlam` and `.xla`.

If you just want to make improvements to the interface, you'll likely not need
to do any of the above. Instead, install from the [automated
installer](https://tulipcell.org). Then open Excel, and open the VBA editor
(Alt-F11). From there you can edit the VBA code, which is where the Excel
add-in features live. If you make useful changes, just send me your new VBA
code. I can work it into the upstream build process.
