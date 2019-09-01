Excel-DNA XFunctions Add-in
===========================

**ExcelDna.XFunctions.xll** is a small add-in that implements two user-defined functions - **[XLOOKUP](https://support.office.com/en-us/article/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929?ui=en-US&rs=en-US&ad=US)** and **[XMATCH](https://support.office.com/en-us/article/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312?ui=en-US&rs=en-US&ad=US)** - that are compatible with the [newly announced built-in functions](https://techcommunity.microsoft.com/t5/Excel-Blog/Announcing-XLOOKUP/ba-p/811376).

For some great material (including videos) on the new functions see the Bill Jelen (Mr. Excel) site - [The VLOOKUP Slayer: XLOOKUP Debuts in Excel]( https://www.mrexcel.com/excel-tips/the-vlookup-slayer-xlookup-debuts-excel/#readmore).

**XFunctions** is meant to be a completely compatible implementation that covers the full functionality of XLOOKUP and XMATCH, but for the current version I expect bugs, especially in the advanced parameter cases.

The add-in should work in all Windows versions of Excel, with separate 32-bit and 64-bit add-ins.

Getting Started
---------------
Binary releases are hosted on GitHub: https://github.com/Excel-DNA/XFunctions/releases

Here are some example workbooks with data from the online help and other blogs posts showing the new functions: https://github.com/Excel-DNA/XFunctions/tree/master/Examples

Examples
--------
The [HelpExamples workbook](https://github.com/Excel-DNA/XFunctions/blob/master/Examples/HelpExamples.xlsx) contains a number of examples corresponding to the online help documentation for the respective functions.

![XLOOKUP Example 1](https://github.com/Excel-DNA/XFunctions/blob/master/Screenshots/XLOOKUPExample1.png)

![XLOOKUP Example 3](https://github.com/Excel-DNA/XFunctions/blob/master/Screenshots/XLOOKUPExample3.png)

![XMATCH Example 1](https://github.com/Excel-DNA/XFunctions/blob/master/Screenshots/XMATCHExample1.png)

Notes
-----
* I've not seen the real `XLOOKUP` or `XMATCH` functions myself, so haven't been able to compare this implementation.
* If you try to debug and see a "Managed Debugger Assistant" message relating to the "LoaderLock" just ignore it... (it is caused by the Excel-DNA IntelliSense extension and should not be a concern)

TODO
----
* Decide whether to use the same names or different ones (e.g. "XLookup_" to make clear it's not the built-in "XLOOKUP" function)
* Understand compatibility for sheets between real functions and the XFunctions version - internally the workbook knows whether a function in a formula is a built-in function or an xll function . . . how does it behave when loaded backwards or forwards?
* Decide what to do when loading in an instance where the real functions are present
* Clean up, add tests etc.

Support and participation
-------------------------
Any help or feedback is greatly appreciated.

"We accept pull requests" ;-) 

Please log bugs and feature suggestions on the GitHub 'Issues' page.

For general comments or discussion, use the Excel-DNA forum at https://groups.google.com/forum/#!forum/exceldna .

License
-------
This project is published under the standard MIT license.

  Govert van Drimmelen
  
  govert@icon.co.za
  
  31 August 2019
