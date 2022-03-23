Option Explicit

Public Sub GetTableDataFromWeb()
    Dim objChrome As Selenium.ChromeDriver
    Dim objTRs As Selenium.WebElements
    Dim objTR As Selenium.WebElement
    Dim objTDs As Selenium.WebElements
    Dim objTD As Selenium.WebElement
    Dim lngRowCount As Long, lngColumnCount As Long, lngRow As Long, r As Long, c As Long, lngLastRow As Long
    Dim Data()
    Const strURL As String = "https://www.worldometers.info/coronavirus/"
    Set objChrome = New Selenium.ChromeDriver
    With objChrome
        .AddArgument "--headless"
        .Start
        .Get strURL
        Set objTRs = .FindElementById("main_table_countries_today").FindElementByTag("tbody").FindElementsByTag("tr")
    End With
    lngRowCount = objTRs.Count
    lngColumnCount = objTRs(1).FindElementsByTag("td").Count
    ReDim Preserve Data(1 To lngRowCount, 1 To lngColumnCount)
    r = 1
    For Each objTR In objTRs
        c = 1
        Set objTDs = objTR.FindElementsByTag("td")
        For Each objTD In objTDs
            Data(r, c) = objTD.Text
            c = c + 1
        Next
        r = r + 1
    Next
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(lngLastRow + 1, 1), Cells(lngLastRow + lngRowCount, lngColumnCount)).Value = Data
    Set objTRs = Nothing
    Set objTR = Nothing
    Set objTDs = Nothing
    Set objTD = Nothing
End Sub
