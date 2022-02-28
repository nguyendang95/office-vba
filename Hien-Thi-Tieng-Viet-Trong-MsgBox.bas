'ThisOutlookSession
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
#End If

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim objItem As Outlook.MailItem
    If TypeOf Item Is Outlook.MailItem Then
        Set objItem = Item
        If CancelNoAttachments(objItem) Then Cancel = True
    End If
    Set objItem = Nothing
End Sub

Private Function CancelNoAttachments(ByVal objItem As Outlook.MailItem) As Boolean
    Dim strMsg As String
    Dim strMsgSet As String
    Dim strKeyword1 As String
    Dim strKeyword2 As String
    Dim strPath As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim fso As Object
    Dim fsoFile As Object
    strPath = "C:\Users\nguye\OneDrive\ThucHanhVBA\CongViec\Outlook-msgbox.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFile = fso.OpenTextFile(strPath, 1, False, -1)
    strMsgSet = fsoFile.ReadAll
    fsoFile.Close
    If objItem.Attachments.Count = 0 Then
        strKeyword1 = ParseTextLinePair(strMsgSet, "attached:")
        strKeyword2 = ParseTextLinePair(strMsgSet, "Attached:")
        intPos1 = InStr(1, objItem.Body, strKeyword1)
        intPos2 = InStr(1, objItem.Body, strKeyword2)
        If intPos1 > 0 Or intPos2 > 0 Then
            strMsg = ParseTextLinePair(strMsgSet, "Check for attachments:")
            If MsgBoxW(strMsg, vbQuestion + vbYesNo, "Add attachments?") = vbYes Then
                CancelNoAttachments = True
            Else: CancelNoAttachments = False
            End If
        End If
    End If
    Set fso = Nothing
    Set fsoFile = Nothing
End Function

'Tùy bi?n hàm MsgBox đ? h? tr? hi?n th? k? t? Unicode
Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft Outlook") As VbMsgBoxResult
Select Case Buttons
    Case vbInformation
        MessageBeep (&H10)
    Case vbQuestion
        MessageBeep (&H20)
    Case vbExclamation
        MessageBeep (&H30)
    Case vbCritical
        MessageBeep (&H40)
    Case Else
        MessageBeep (&H0)
End Select
MsgBoxW = MessageBoxW(GetFocus(), StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

Public Function ParseTextLinePair(strSource As String, strLabel As String)
    Dim intLocLabel As Integer
    Dim intLocCRLF As Integer
    Dim intLenLabel As Integer
    Dim strText As String
    'Lay vi tri cua chuoi ky tu label trong van ban nguon
    intLocLabel = InStr(1, strSource, strLabel)
    'Tinh do dai chuoi ky tu label
    intLenLabel = Len(strLabel)
    'Neu ton tai chuoi ky tu label thi thuc hien buoc tiep theo
    If intLocLabel > 0 Then
        'Tim vi tri ky tu xuong dong, bat dau tu vi tri chuoi ky tu label
        intLocCRLF = InStr(intLocLabel, strSource, vbCrLf)
        'Tien hanh tach chuoi label
        If intLocCRLF > 0 Then
            intLocLabel = intLocLabel + intLenLabel
            strText = Mid(strSource, intLocLabel, intLocCRLF - intLocLabel)
        Else: strText = Mid(strSource, intLocLabel + intLenLabel)
        End If
    End If
    ParseTextLinePair = Trim(strText)
End Function
