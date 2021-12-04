Option Explicit
'Macro này tự động trộn các câu hỏi trắc nghiệm từ bộ đề thi có sẵn để tạo thành bộ đề trắc nghiệm mới. 
'Cần tạo bookmark cho các câu hỏi trắc nghiệm trước khi chạy macro,
Sub TaoDeThiMoiTuKhoDeChoSan()
    Dim objDoc As Document
    Dim objNewDoc As Document
    Dim objBkmk As Bookmark
    Dim numrandom As Long
    Dim SoLuong As Long
    Dim WorkProcessed As Long
    Dim objDic As Object
    Dim objFldl As FileDialog
            
    WorkProcessed = 0
    Set objFldl = Application.FileDialog(msoFileDialogOpen)
    With objFldl
        .AllowMultiSelect = False
        .Title = "Select a test document"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx; *.doc", 1
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            Set objDoc = Documents.Open(FileName:=.SelectedItems(1), ReadOnly:=True)
        End If
    End With
    If objDoc.Bookmarks.Count = 0 Then
        MsgBoxW "Tap tin ban chon khong chua noi dung de thi trac nghiem hoac ban chua thiet lap bookmark cho cac cau hoi trac nghiem.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    SoLuong = InputBox("Ban muon chon ngau nhien bao nhieu cau hoi trac nghiem de tao ra bo de thi moi?")
    If SoLuong < 0 Then
        MsgBoxW "So luong cau hoi trac nghiem phai lon hon 0!", vbExclamation
        Exit Sub
    End If
    If Err.Number <> 0 Then
        MsgBoxW "Vui long chi nhap so nguyen, khong am, khong chua bat ky ky tu khong phai so.", vbExclamation
        Exit Sub
    End If
    Set objNewDoc = Documents.Add
    objDoc.Activate
    Set objDic = CreateObject("Scripting.Dictionary")
    Do Until WorkProcessed = SoLuong
        numrandom = Int((SoLuong - 1 + 1) * Rnd + 1)
        If Not objDic.Exists(numrandom) Then
            objDic.Add numrandom, ""
            Set objBkmk = objDoc.Bookmarks.Item(numrandom)
            objBkmk.Range.Copy
            objNewDoc.Activate
            With Selection
                .Paste
                .TypeParagraph
            End With
            objDoc.Activate
            WorkProcessed = WorkProcessed + 1
        End If
    Loop
    objDoc.Quit
    objNewDoc.SaveAs2 FileName:=objDoc.Path & "\" & "DeThiTracNghiemTiengAnhDaTron.docx"
    Set objDoc = Nothing
    Set objNewDoc = Nothing
    Set objBkmk = Nothing
    Set objDic = Nothing
    Set objFldl = Nothing
End Sub
