Option Explicit

Public Function RegexExtract(Range As Excel.Range, Pattern As String) As String
    Dim objRegex As Object
    Dim objRegexMatch As Object
    Dim colRegexMatches As Object
    Set objRegex = CreateObject("VBScript.Regexp")
    With objRegex
        .Pattern = Pattern
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
    End With
    Set colRegexMatches = objRegex.Execute(Range.Value)
    For Each objRegexMatch In colRegexMatches
        RegexExtract = RegexExtract & " " & objRegexMatch.Value
    Next
    Set objRegex = Nothing
    Set objRegexMatch = Nothing
    Set colRegexMatches = Nothing
End Function

Public Function RegexReplace(Range As Excel.Range, Pattern As String, ReplacementString As String) As String
    Dim objRegex As Object
    Dim Result As String
    Set objRegex = CreateObject("VBScript.Regexp")
    With objRegex
        .Pattern = Pattern
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
    End With
    Result = objRegex.Replace(Range.Value, ReplacementString)
    RegexReplace = Result
    Set objRegex = Nothing
End Function

Public Function RegexMatch(Range As Excel.Range, Pattern As String) As Boolean
    Dim objRegex As Object
    Dim colMatches As Object
    Set objRegex = CreateObject("VBScript.Regexp")
    With objRegex
        .Pattern = Pattern
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
    End With
    Set colMatches = objRegex.Execute(Range.Value)
    If colMatches.Count > 0 Then
        RegexMatch = True
    Else: RegexMatch = False
    End If
    Set objRegex = Nothing
    Set colMatches = Nothing
End Function
