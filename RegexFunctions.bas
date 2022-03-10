Attribute VB_Name = "RegexFunctions"
' REFERENCE "Microsoft VBScript Regular Expressions 5.5" FOR ( RegExp )
Function RegexReplace(pat As String, _
    repl As String, _
    str As String, _
    Optional is_global As Boolean = True, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As String
' Replace all instances of pat in str with repl.
Dim regex As Object
Set regex = New RegExp
With regex:
    .Global = is_global
    .IgnoreCase = ignore_case
    .Pattern = pat
    .multiline = multiline
End With
RegexReplace = regex.Replace(str, repl)
End Function


Function RegexContains(pat As String, _
    str As String, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As Boolean
' Return True if regular expression pat matches the string else False
Dim regex: Set regex = New RegExp
With regex:
    .IgnoreCase = ignore_case
    .Pattern = pat
    .multiline = multiline
End With
RegexContains = regex.Test(str)
End Function


Function RegexFullMatch(pat As String, _
    str As String, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As Boolean
' Return True if regular expression pat matches the string EXACTLY else False
Dim regex: Set regex = New RegExp
With regex:
    .IgnoreCase = ignore_case
    .Pattern = "^" + pat + "$"
    .multiline = multiline
End With
RegexFullMatch = regex.Test(str)
End Function


Function RegexMatch(pat As String, _
    str As String, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As Boolean
' Return True if regular expression pat matches the string BEGINNING else False
Dim regex: Set regex = New RegExp
With regex:
    .IgnoreCase = ignore_case
    .Pattern = "^" + pat
    .multiline = multiline
End With
RegexMatch = regex.Test(str)
End Function


Function RegexMatchEnd(pat As String, _
    str As String, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As Boolean
' Return True if regular expression pat matches the string END else False
Dim regex: Set regex = New RegExp
With regex:
    .IgnoreCase = ignore_case
    .Pattern = pat + "$"
    .multiline = multiline
End With
RegexMatchEnd = regex.Test(str)
End Function


Function RegexMatches(pat As String, _
    str As String, _
    Optional is_global As Boolean = True, _
    Optional ignore_case As Boolean = True, _
    Optional multiline As Boolean = False) As Object
' Get all the matches for a pattern in string str with those parameters
Dim regex: Set regex = New RegExp
With regex:
    .Global = is_global
    .IgnoreCase = ignore_case
    .Pattern = pat
    .multiline = multiline
End With
Set RegexMatches = regex.Execute(str) ' find all matches
End Function


Function RegexFindAll(pat As String, _
    str As String, _
    Optional sep As String = ", ", _
    Optional is_global As Boolean = True, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As String
' Find all matches and stringjoin them together with sep.
' Unlike RegexMatches, this is suitable for use as a worksheet formula.
Dim matches
Dim num_matches As Integer
Dim ii As Integer
Set matches = RegexMatches(pat, str, is_global, ignore_case, multiline)
num_matches = matches.Count - 1
ReDim strings(num_matches) As String ' this will be joined at the end
For ii = 0 To num_matches
    ' include every match found, with sep (the argument) separating
    ' each match
    strings(ii) = matches.Item(ii)
Next ii
RegexFindAll = Join(strings, sep)
End Function


Function RegexFind(pat As String, _
    str As String, _
    Optional match_num As Integer = 0, _
    Optional submatch_sep As String = "", _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False) As String
' Get the match_num^th match in string str to a regex with various params.
Dim is_global As Boolean
Dim matches
Dim num_submatches As Integer
is_global = IIf(match_num = 0, False, True)
Set matches = RegexMatches(pat, str, is_global, ignore_case, multiline)
If matches.Count - 1 < match_num Then
    MsgBox ("Match number " + match_num _
    + " couldn't be found in matches of length " + matches.Count)
    Exit Function
End If
num_submatches = matches.Item(match_num).SubMatches.Count - 1
If num_submatches < 0 Then
    RegexFind = matches.Item(match_num).Value
Else
    ReDim strings(num_submatches) As String
    Dim ii As Integer
    For ii = 0 To num_submatches
        strings(ii) = matches.Item(match_num).SubMatches.Item(ii)
    Next ii
    RegexFind = Join(strings, submatch_sep)
End If
End Function


Function RegexSplit(pat As String, _
    str As String, _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False, _
    Optional num_splits = -1) As Variant
' If there are no capturing groups in pat, get an array of all the substrings not matched
' by the regex.
' if there are capturing groups in the regex, each capturing group gets its
' own element in the string list at the location where it's found in the string
' in addition to all the substrings not matched by the regex
Dim is_global As Boolean
Dim matches
is_global = IIf(num_splits = 0, False, True)
Set matches = RegexMatches(pat, str, is_global, ignore_case, multiline)
ReDim out(1) As Variant
If matches.Count = 0 Then
    out(0) = str
Else
    Dim num_strings As Integer
    Dim matches_so_far As Integer
    Dim first_index As Long
    Dim match_index As Long
    Dim unmatched As String
    Dim num_submatches As Integer
    Dim submatch
    num_submatches = matches.Item(0).SubMatches.Count
    ReDim out(matches.Count * (num_submatches + 1))
    For Each match In matches
        If (num_splits > 0) And (matches_so_far = num_splits) Then Exit For
        match_index = match.FirstIndex
        unmatched = Mid(str, first_index + 1, match_index - first_index)
        ' VBA's Mid function uses 1-based indexing
        first_index = match_index + match.Length
        out(num_strings) = unmatched
        num_strings = num_strings + 1
        For Each submatch In match.SubMatches
            out(num_strings) = submatch
            num_strings = num_strings + 1
        Next submatch
        matches_so_far = matches_so_far + 1
        ' MsgBox ("{" + Join(out, ", ") + "}")
    Next match
    out(num_strings) = Mid(str, first_index + 1, Len(str) - first_index)
    ReDim Preserve out(num_strings)
End If
RegexSplit = out
End Function


Function RegexSplitToString(pat As String, _
    str As String, _
    Optional sep As String = ", ", _
    Optional ignore_case As Boolean = False, _
    Optional multiline As Boolean = False, _
    Optional num_splits = -1) As String
' Uses RegexSplit to split out all instances of the regex
' (or split and include capturing groups as described above)
' and then stringjoin the resulting array with sep.
' Unlike RegexSplit, this is suitable for use as a worksheet formula.
Dim strings As Variant
strings = RegexSplit(pat, str, ignore_case, multiline, num_splits)
RegexSplitToString = Join(strings, sep)
End Function

