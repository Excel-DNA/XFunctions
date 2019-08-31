Excel-DNA XFunctions Add-in
===========================

**ExcelDna.XFunctions.xll** is a small add-in that implements two user-defined functions - **[XLOOKUP](https://support.office.com/en-us/article/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929?ui=en-US&rs=en-US&ad=US)** and **[XMATCH](https://support.office.com/en-us/article/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312?ui=en-US&rs=en-US&ad=US)** - that are compatible with the [newly announced built-in functions](https://techcommunity.microsoft.com/t5/Excel-Blog/Announcing-XLOOKUP/ba-p/811376).

So far I have only implemented the default exact first-item match case (with and without wildcards), but for that case the two functions should work correctly, even returning a reference when the built-in XLOOKUP function would do so.

The add-in should work in all Windows versions of Excel, with separate 32-bit and 64-bit add-ins.

Notes
-----
* I've not seen the real `XLOOKUP` or `XMATCH` functions myself, so haven't been able to compare this implementation.
* I've not released binaries - I think we should cover all the parameters first
* If you try to debug and see a "Managed Debugger Assistant" message relating to the "LoaderLock" just ignore it... (it is caused by the Excel-DNA IntelliSense extension and should not be a concern)

TODO
----
* Decide whether to use the same names or different ones (e.g. "XLookup_" to make clear it's not the built-in "XLOOKUP" function)
* Understand compatibility for sheets between real functions and the XFunctions version - internally the workbook knows whether a function in a formula is a built-in function or an xll function . . . how does it behave when loaded backwards or forwards?
* Decide what to do when loading in an instance where the real functions are present
* Finish the XMATCH implementation for different switches
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
