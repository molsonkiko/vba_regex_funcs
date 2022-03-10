# Change Log
All notable changes to this project will be documented in this file.
 
The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).
 
## [Unreleased] - yyyy-mm-dd
 
### To Be Added

- ???

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