Option Explicit
Private WithEvents colItems As Outlook.Items

Private Sub Class_Initialize()
    Dim objInboxFldr As Outlook.Folder
    Set objInboxFldr = Application.Session.DefaultStore.GetDefaultFolder(olFolderInbox)
    Set colItems = objInboxFldr.Items
End Sub

Private Sub colItems_ItemAdd(ByVal Item As Object)
    Dim objMail As Outlook.MailItem
    Dim objXlApp As Object
    Dim objWb As Object
    Dim objEmailSh As Object
    Dim objPivotSh As Object
    Dim objTable As Object
    Dim objListRows As Object
    Dim objPivotTable As Object
    Dim arrData()
    Const strWbPath = "C:\Users\nguye\OneDrive\ThucHanhVBA\ThongKeEmail.xlsm"
    If TypeOf Item Is Outlook.MailItem Then
        Set objMail = Item
        With objMail
            arrData = Array(.Subject, .SenderName, .SenderEmailAddress, .ReceivedTime)
        End With
        Set objXlApp = CreateObject("Excel.Application")
        Set objWb = objXlApp.Workbooks.Open(strWbPath)
        Set objEmailSh = objWb.Sheets("EmailDetails")
        Set objTable = objEmailSh.ListObjects(1)
        Set objListRows = objTable.ListRows.Add
        With objListRows
            .Range(1) = arrData(0)
            .Range(2) = arrData(1)
            .Range(3) = arrData(2)
            .Range(4) = arrData(3)
        End With
    End If
    Set objPivotSh = objWb.Sheets("EmailPivotTable")
    Set objPivotTable = objPivotSh.PivotTables(1)
    objPivotTable.Refreshtable
    objWb.Close SaveChanges:=True
    objXlApp.Quit
    Set objMail = Nothing
    Set objXlApp = Nothing
    Set objWb = Nothing
    Set objEmailSh = Nothing
    Set objPivotSh = Nothing
    Set objTable = Nothing
    Set objListRows = Nothing
    Set objPivotTable = Nothing
End Sub
