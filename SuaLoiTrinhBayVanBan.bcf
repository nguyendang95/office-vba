Option Explicit

Sub SuaLoiVanBan()
Dim i As Byte

Selection.HomeKey wdStory, wdMove
'Xoa khoang trong du thua
With Selection.Find
    .ClearFormatting
    .Forward = True
    .Text = "  "
    .Wrap = wdFindContinue
    .Replacement.Text = " "
    .Execute Replace:=wdReplaceAll
End With
'Xoa khoang trong thua truoc dau cham cau
With Selection.Find
    .ClearFormatting
    .Forward = True
    .Text = " ."
    .Wrap = wdFindContinue
    .Replacement.Text = "."
    .Execute Replace:=wdReplaceAll
End With
'Xoa khoang trong thua truoc dau phay
With Selection.Find
    .Forward = True
    .MatchWildcards = True
    .Text = " ,"
    .Wrap = wdFindContinue
    .Replacement.Text = ", "
    .Execute Replace:=wdReplaceAll
End With
'Viet hoa truoc dau cham
With Selection.Find
    .ClearFormatting
    .Forward = True
    .MatchWildcards = True
    .Text = ". [a-z]"
    .Execute
    Do
        Selection.Range.Case = wdUpperCase
        Selection.Collapse wdCollapseEnd
        .Execute
    Loop While .Found
End With
'Viet hoa dau cau
For i = 65 To 122
With Selection.Find
    .ClearFormatting
    .Forward = True
    .Text = Chr(13) & Chr(i - 32)
    .Wrap = wdFindContinue
    .Replacement.Text = Chr(13) & Chr(i)
    .Execute Replace:=wdReplaceAll
End With

End Sub
