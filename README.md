VBA regex functions
============

*VBA adapations of Python re functions*

Features
--------

* Regular expression functions adapted from python's [re](https://docs.python.org/3/library/re.html) library.
* Several functions that can be used from a worksheet once the add-in is enabled.

How to use
----------

* The functions can be used in VBA macros if you want, but if you want to use them as worksheet functions, do the following:
	1. Open Excel.
	2. Go to Options -> Add-ins -> Go -> Browse
	3. Select the path to RegexFunctions.xlam in your file explorer
	4. Open up regexfunctions_addin_test.xlsx and use F9 to refresh all the functions. You can visually inspect the results and source code to make sure it looks like successful_test_outcomes.png.
* To use these functions in a VBA Macro, do the following:
	1. Open Excel.
	2. Use Alt-F11 to open macros.
	3. Go to Tools -> References -> Browse and change the type of extensions you're looking for to "Microsoft Excel Files".
	4. Select the path to RegexFunctions.xlam in your file explorer.
	5. While you're at it, search for "Microsoft VBScript Regular Expressions 5.5" on the checklist and check that. This will allow you to use RegExp objects. The references are ordered alphabetically.
* See the CHANGELOG for what functions you can use.
* Also see [regular-expressions.info](https://www.regular-expressions.info/vbscript.html) for more information on the object model that VBA regular RegExp objects use.

Contributing
------------

Be sure to read the [contribution guidelines]

(https://github.com/molsonkiko/vba_regex_funcs/blob/main/CONTRIBUTING.md).