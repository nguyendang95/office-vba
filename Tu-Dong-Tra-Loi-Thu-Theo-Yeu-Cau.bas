Option Explicit

Private WithEvents colItems As Outlook.Items

Private Sub Class_Initialize()
    Dim objFolder As Outlook.Folder
    Set objFolder = Application.Session.GetDefaultFolder(olFolderInbox)
    Set colItems = objFolder.Items
End Sub

Private Sub colItems_ItemAdd(ByVal Item As Object)
    Dim objMail As Outlook.MailItem
    Dim objReplyMail As Outlook.MailItem
    Dim objXlApp As Object
    Dim objXlWb As Object
    Dim objXlDataSh As Object
    Dim objXlResultSh As Object
    Dim objXlTable As Object
    Dim objDataRng As Object
    Dim objXlListRow As Object
    Dim OldTableRowCount As Long, i As Long, j As Long, tblLastRow As Long
    Dim strFileName As String
    Dim Data()
    Dim objRegEx As Object
    Dim objMatches As Object, objMatch As Object
    If Item.Class = olMail Then
        Set objMail = Item
        Set objXlApp = CreateObject("Excel.Application")
        'Ten tap tin Excel can lay du lieu
        Const strWbPath = Environ$("USERPROFILE") & "\Documents\template-excel.xlsx"
        Set objXlWb = objXlApp.Workbooks.Open(strWbPath, , True)
        'Ten Sheet chua du lieu can tuong tac
        Set objXlDataSh = objXlApp.Sheets("2016")
        If objXlDataSh.AutoFilterMode Then objXlDataSh.AutoFilterMode = False
        'Ten Sheet chua du lieu sau khi da loc AutoFilter
        Set objXlResultSh = objXlApp.Sheets("TableResult")
        Set objXlTable = objXlResultSh.ListObjects(1)
        OldTableRowCount = objXlTable.ListRows.Count
        For i = OldTableRowCount To 1 Step -1
            objXlTable.ListRows.Item(i).Delete
        Next
        Set objDataRng = objXlDataSh.Range("A1").CurrentRegion
        Set objRegEx = CreateObject("VBScript.RegExp")
        With objRegEx
            .Pattern = "CDISC-[0-9]{4}"
            .Global = True
        End With
        Set objMatches = objRegEx.Execute(objMail.body)
        tblLastRow = objXlTable.Range.Rows.Count
        Const xlCellTypeVisible As Byte = 12
        For Each objMatch In objMatches
            objDataRng.AutoFilter 3, objMatch.Value
            Data() = objDataRng.Offset(1).Resize(objDataRng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Value
            Set objXlListRow = objXlTable.ListRows.Add(tblLastRow - 1)
            objXlListRow.Range.Value = Data()
        Next
        Const xlTypePDF As Byte = 0
        'Thu muc xuat ra ket qua, tap tin PDF                                             
        Const strExportFolder = Environ$("USERPROFILE") & "\Documents"
        strFileName = strExportFolder & "\RequestCodesResult.pdf"
        If Dir(strFileName) <> vbNullString Then Kill strFileName
        objXlTable.Range.ExportAsFixedFormat xlTypePDF, strFileName
        objXlWb.Close SaveChanges:=False
        objXlApp.Quit
        Set objReplyMail = objMail.Reply
        With objReplyMail
            .Display
            .BodyFormat = olFormatHTML
            .HTMLBody = "<p>Hello,</p><br>" & _
                    "<p>Here is the result.</p>"
            .Attachments.Add strFileName
            '.Send /.Send de gui thu ngay lap tuc
        End With
    End If
    Set objMail = Nothing
    Set objReplyMail = Nothing
    Set objAtt = Nothing
    Set objXlDataSh = Nothing
    Set objXlResultSh = Nothing
    Set objXlTable = Nothing
    Set objDataRng = Nothing
    Set objXlListRow = Nothing
    Set objRegEx = Nothing
    Set objMatches = Nothing
    Set objMatch = Nothing
End Sub
