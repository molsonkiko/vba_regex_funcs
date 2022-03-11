# Change Log
All notable changes to this project will be documented in this file.
 
The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).
 
## [Unreleased] - yyyy-mm-dd
 
### To Be Added

- Support for lazy quanitifiers in the Regex Builder Form.

## [2.1.0] - 2022-03-10

### Added

- The "Regex Builder Form" is complete! With this, you can build complex regular expressions from simpler ones using a GUI, and the worksheet runs all of the functions in RegexFunctions on an input string of your choice (default "a big bad dog") and the last regex you created.
- Currently the Regex Builder Form does not allow you to create regexes with [lookahead](https://www.regular-expressions.info/lookaround.html#lookahead) or [lazy quantifiers](https://www.regular-expressions.info/repeat.html#lazy).

## [2.0.0] - 2022-03-10

### Added

- RegexFindObject (like indexing in `list(re.finditer(x))`, which returns the n^th match as a match object, unlike RegexFind, which returns a string.
- RegexEscape (`re.escape` workalike), which takes a string and returns the string with all special characters escaped.
- The basic skeleton of "Regex Builder Form.xlsm", which will allow people to build regular expressions by a GUI similar to [RegexBuddy](https://www.rexegg.com/regexbuddy-tutorial.html). This is not yet functional.

## [1.0.0] - 2022-03-09

### Added

- RegexFunctions.bas, which contains the source code in a form you can read on your favorite text editor
- RegexFunctions.xlam, an Excel add-in that requires Microsoft VBScript Regular Expressions 5.5
- regexfunctions_addin_test.xlsx, which contains tests for the worksheet functions from RegexFunctions.
- These functions are included (corresponding Python `re` functions in parens):
	- RegexReplace (workalike of `re.sub`, except no function replacements)
	- RegexContains (`bool(re.search(x))`)
	- RegexFullMatch (`bool(re.fullmatch(x))`)
	- RegexMatch (`bool(re.match(x))`)
	- RegexMatchEnd (like `bool(re.match(x))`, but matches the end)
	- RegexMatches (`re.finditer`)
	- RegexFindall (sort of like stringjoining `re.findall`)
	- RegexFind (sort of like indexing in the list from `re.findall`)
	- RegexSplit(like `re.split`)
	- RegexSplitToString (like stringjoining the list from `re.split`)
- All of those functions except RegexSplit and RegexMatches can be used as worksheet functions once the add-in has been enabled.