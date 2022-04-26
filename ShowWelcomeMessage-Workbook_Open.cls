'Class Module 'Welcome'
Option Explicit

Private WithEvents objXlApp As Excel.Application

Private Sub Class_Initialize()
    Set objXlApp = Application
End Sub

Private Sub Class_Terminate()
    Set objXlApp = Nothing
End Sub

Private Sub objXlApp_WorkbookOpen(ByVal Wb As Workbook)
    MsgBox "Welcome to Excel", vbInformation, "Welcome"
End Sub
