Option Explicit
Public Sub CreateFileList()
        Dim objWb As Excel.Workbook
        Dim objSh As Excel.Worksheet
        Dim objFSO, objFile, objFolder, objSubFolder, objSubFolders As Object
        Dim objFolderPicker As FileDialog
        Dim lngRow, lngCount As Long
        Dim strPath As String
        Set objFolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With objFolderPicker
            .Title = "Select a folder"
            .Show
            If .SelectedItems.Count = 0 Then
                Exit Sub
            Else: strPath = .SelectedItems(1)
            End If
        End With
        Application.ScreenUpdating = False
        Set objWb = Application.Workbooks.Add
        Set objSh = objWb.Sheets(1)
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder(strPath)
        Set objSubFolders = objFolder.SubFolders
        With objSh
            .Range("A2").Value = "File List"
            .Range("A2").Font.Size = 16
            .Range("A2").Font.Bold = True
            .Range("A3").Value = "Folder: " & objFolder.Path
            .Range("A3").Offset(1, 0).Value = "Count:"
            .Range("A6").Value = "File Name"
            .Range("A6").Offset(0, 1).Value = "Path"
            .Range("A6").Offset(0, 2).Value = "File Type"
            .Range("A6").Offset(0, 3).Value = "Date Created"
            .Range("A6").Offset(0, 4).Value = "Date Last Accessed"
            .Range("A6").Offset(0, 5).Value = "Date Last Modified"
            .Range("A6").Offset(0, 6).Value = "Size (KB)"
            .Range("A6").Offset(0, 7).Value = "Link"
        End With
        lngRow = 7
        For Each objSubFolder In objSubFolders
            For Each objFile In objSubFolder.Files
                With objFile
                    lngCount = lngCount + 1
                    Debug.Print .Name
                    objSh.Cells(lngRow, 1).Value = .Name
                    objSh.Cells(lngRow, 2).Value = .ParentFolder
                    objSh.Cells(lngRow, 4).Value = .DateCreated
                    objSh.Cells(lngRow, 5).Value = .DateLastAccessed
                    objSh.Cells(lngRow, 6).Value = .DateLastModified
                    objSh.Cells(lngRow, 7).Value = BytesToKilobytes(.Size)
                    objSh.Hyperlinks.Add Anchor:=objSh.Cells(lngRow, 8), Address:=.Path, ScreenTip:="Click to open", TextToDisplay:="Open"
                    Select Case GetFileExtension(.Name)
                        Case ".xlsx": objSh.Cells(lngRow, 3).Value = "Excel Workbook"
                        Case ".xls": objSh.Cells(lngRow, 3).Value = "Excel 97-2003 Workbook"
                        Case ".xlsm": objSh.Cells(lngRow, 3).Value = "Excel Macro-Enabled Workbook"
                        Case ".xlam": objSh.Cells(lngRow, 3).Value = "Excel Add-In"
                        Case ".xla": objSh.Cells(lngRow, 3).Value = "Excel 97-2003 Add-In"
                        Case ".accdb": objSh.Cells(lngRow, 3).Value = "Microsoft Access database file"
                        Case ".accdt": objSh.Cells(lngRow, 3).Value = "Microsoft Access database template"
                        Case ".csv": objSh.Cells(lngRow, 3).Value = "Comma-separated value file"
                        Case ".doc": objSh.Cells(lngRow, 3).Value = "Microsoft Word document before Word 2007"
                        Case ".docx": objSh.Cells(lngRow, 3).Value = "Microsoft Word document"
                        Case ".docm": objSh.Cells(lngRow, 3).Value = "Microsoft Word macro-enabled document"
                        Case ".dot": objSh.Cells(lngRow, 3).Value = "Microsoft Word template before Word 2007"
                        Case ".dotx": objSh.Cells(lngRow, 3).Value = "Microsoft Word template"
                        Case ".flv": objSh.Cells(lngRow, 3).Value = "Flash-compatible video file"
                        Case ".gif": objSh.Cells(lngRow, 3).Value = "Graphical Interchange Format file"
                        Case ".iso": objSh.Cells(lngRow, 3).Value = "ISO-9660 disc image"
                        Case ".jpg", ".jpeg", ".JPG": objSh.Cells(lngRow, 3).Value = "Joint Photographic Experts Group photo file"
                        Case ".m4a": objSh.Cells(lngRow, 3).Value = "MPEG-4 audio file"
                        Case ".mdb": objSh.Cells(lngRow, 3).Value = "Microsoft Access database before Access 2007"
                        Case ".mid", ".midi": objSh.Cells(lngRow, 3).Value = "Musical Instrument Digital Interface file"
                        Case ".mov": objSh.Cells(lngRow, 3).Value = "Apple QuickTime movie file"
                        Case ".mp3": objSh.Cells(lngRow, 3).Value = "MPEG layer 3 audio file"
                        Case ".mp4": objSh.Cells(lngRow, 3).Value = "MPEG 4 video"
                        Case ".mpeg": objSh.Cells(lngRow, 3).Value = "Moving Picture Experts Group movie file"
                        Case ".msi": objSh.Cells(lngRow, 3).Value = "Window installer file"
                        Case ".pdf": objSh.Cells(lngRow, 3).Value = "Portable Document Format file"
                        Case ".png": objSh.Cells(lngRow, 3).Value = "Portable Network Graphics file"
                        Case ".pot": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint template before Powerpoint 2007"
                        Case ".potm": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint macro-enabled template"
                        Case ".potx": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint template"
                        Case ".ppam": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint add-in"
                        Case ".pps": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint slideshow before Powerpoint 2007"
                        Case ".ppsm": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint macro-enabled slideshow"
                        Case ".ppsx": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint slideshow"
                        Case ".ppt": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint format before Powerpoint 2007"
                        Case ".pptm": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint macro-enabled presentation"
                        Case ".pptx": objSh.Cells(lngRow, 3).Value = "Micrososft Powerpoint presentation"
                        Case ".psd": objSh.Cells(lngRow, 3).Value = "Adobe Photoshop file"
                        Case ".pst": objSh.Cells(lngRow, 3).Value = "Outlook data store"
                        Case ".pub": objSh.Cells(lngRow, 3).Value = "Microsoft Publisher file"
                        Case ".rar": objSh.Cells(lngRow, 3).Value = "Roshal Archive compressed file"
                        Case ".rft": objSh.Cells(lngRow, 3).Value = "Rich Text Format file"
                        Case ".sldm": objSh.Cells(lngRow, 3).Value = "Microsoft Powerpoint macro-enabled slide"
                        Case ".swf": objSh.Cells(lngRow, 3).Value = "Shockware Flash file"
                        Case ".tif", ".tiff": objSh.Cells(lngRow, 3).Value = "Tagged Image Format file"
                        Case ".txt": objSh.Cells(lngRow, 3).Value = "Unformatted text file"
                        Case ".vob": objSh.Cells(lngRow, 3).Value = "Video object file"
                        Case ".vsd": objSh.Cells(lngRow, 3).Value = "Microsoft Visio drawing before Visio 2013"
                        Case ".vsdm": objSh.Cells(lngRow, 3).Value = "Microsoft Visio macro-enabled drawing"
                        Case ".vsdx": objSh.Cells(lngRow, 3).Value = "Microsoft Visio drawing file"
                        Case ".vss": objSh.Cells(lngRow, 3).Value = "Microsoft Visio stencil before Visio 2013"
                        Case ".vssm": objSh.Cells(lngRow, 3).Value = "Microsoft Visio macro-enabled stencil"
                        Case ".vst": objSh.Cells(lngRow, 3).Value = "Microsoft Visio template before Visio 2013"
                        Case ".vstm": objSh.Cells(lngRow, 3).Value = "Microsoft Visio macro-enabled template"
                        Case ".vstx": objSh.Cells(lngRow, 3).Value = "Microsoft Visio template"
                        Case ".wav": objSh.Cells(lngRow, 3).Value = "Wave audio file"
                        Case ".wbk": objSh.Cells(lngRow, 3).Value = "Microsoft Word backup documnet"
                        Case ".wks": objSh.Cells(lngRow, 3).Value = "Microsoft Works file"
                        Case ".wma": objSh.Cells(lngRow, 3).Value = "Windows Media Audio file"
                        Case ".wmd": objSh.Cells(lngRow, 3).Value = "Windows Media Download file"
                        Case ".wmv": objSh.Cells(lngRow, 3).Value = "Windows Media Video File"
                        Case ".wpd", ".wp5": objSh.Cells(lngRow, 3).Value = "WordPerfect document"
                        Case ".xla": objSh.Cells(lngRow, 3).Value = "Microsoft Excel add-in or macro file"
                        Case ".xll": objSh.Cells(lngRow, 3).Value = "Microsoft Excel DLL-based add-in"
                        Case ".xlm": objSh.Cells(lngRow, 3).Value = "Microsoft Excel macro before Excel 2007"
                        Case ".xlt": objSh.Cells(lngRow, 3).Value = "Microsoft Excel template before Excel 2007"
                        Case ".xltm": objSh.Cells(lngRow, 3).Value = "Microsoft Excel macro-enabled after Excel 2007"
                        Case ".xltx": objSh.Cells(lngRow, 3).Value = "Microsoft Excel template after Excel 2007"
                        Case ".xps": objSh.Cells(lngRow, 3).Value = "XML-based document"
                        Case ".zip": objSh.Cells(lngRow, 3).Value = "Compressed file"
                        Case ".htm", ".html": objSh.Cells(lngRow, 3).Value = "Hypertext markup language page"
                        Case ".avi": objSh.Cells(lngRow, 3).Value = "Audio Video Interleave movie or sound file"
                        Case ".bmp": objSh.Cells(lngRow, 3).Value = "Bitmap file"
                        Case ".dll": objSh.Cells(lngRow, 3).Value = "Dynamic Link Library file"
                        Case ".bat": objSh.Cells(lngRow, 3).Value = "PC batch file"
                        Case ".vbs": objSh.Cells(lngRow, 3).Value = "Visual Basic Script file"
                        Case ".aac", ".adt", ".adts": objSh.Cells(lngRow, 3).Value = "Windows audio file"
                        Case ".odt": objSh.Cells(lngRow, 3).Value = "OpenDocument"
                        Case Else: objSh.Cells(lngRow, 3).Value = "Unknown file extension"
                    End Select
                End With
                lngRow = lngRow + 1
            Next
        Next
        objSh.Range("A4").Value = objSh.Range("A4").Value & " " & lngCount
        objSh.Range("A:H").Columns.AutoFit
        objSh.Range("A6:H6").Font.Bold = True
        objSh.Range("A6").CurrentRegion.Borders.LineStyle = xlContinuous
        Application.ScreenUpdating = True
        Set objWb = Nothing
        Set objSh = Nothing
        Set objFSO = Nothing
        Set objFile = Nothing
        Set objFolder = Nothing
        Set objSubFolder = Nothing
        Set objSubFolders = Nothing
        Set objFolderPicker = Nothing
    End Sub
                                                                                                   
    Private Function GetFileExtension(FileName As String) As String
        On Error GoTo BlankFileExtension
        GetFileExtension = Mid(FileName, InStrRev(FileName, "."))
        If Err.Number = 5 Then
BlankFileExtension: GetFileExtension = vbNullString
        End If
    End Function

    Private Function BytesToKilobytes(Bytes As Long) As Long
        BytesToKilobytes = Round(Bytes / 1000, 0)
    End Function
