Private Sub Application_Startup()
    Call ExpandFolders
End Sub

Private Sub ExpandFolders()
    Dim objStore As Outlook.Store
    Dim objCurrentFolder As Outlook.Folder
    Dim objRootFolder As Outlook.Folder
    Dim objFolder As Outlook.Folder
    Set objCurrentFolder = Application.ActiveExplorer.CurrentFolder
    For Each objStore In Application.Session.Stores
        Set objRootFolder = objStore.GetRootFolder
        For Each objFolder In objRootFolder.Folders
            Call LoopFolders(objFolder)
        Next
        Set Application.ActiveExplorer.CurrentFolder = objCurrentFolder
    Next
    Set objStore = Nothing
    Set objCurrentFolder = Nothing
    Set objRootFolder = Nothing
    Set objFolder = Nothing
End Sub

Private Sub LoopFolders(objFolder As Outlook.Folder)
    Dim objSubFolder As Outlook.Folder
    Set Application.ActiveExplorer.CurrentFolder = objFolder
    If objFolder.Folders.Count > 0 Then
        For Each objSubFolder In objFolder.Folders
            Call LoopFolders(objSubFolder)
        Next
    End If
    Set objSubFolder = Nothing
End Sub
