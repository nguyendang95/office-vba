Option Explicit
'Public Sub objChrome As Selenium.ChromeDriver As Selenium.ChromeDriver khai bao bien objChrome de ngan khong cho Chrome tat khi ket thuc thu tuc Sub

Public Sub GetTableDataFromWeb()
    Dim objChrome As Selenium.ChromeDriver 'Khai bao objChrome trong thu tuc se khien cho Chrome bi tat khi ket thuc thu tuc Sub
    Dim objTRs As Selenium.WebElements
    Dim objTR As Selenium.WebElement
    Dim objTDs As Selenium.WebElements
    Dim objTD As Selenium.WebElement
    Dim lngRowCount As Long, lngColumnCount As Long, lngRow As Long, r As Long, c As Long, lngLastRow As Long
    Dim Data()
    Const strURL As String = "https://www.worldometers.info/coronavirus/" 'Dia chi trang web
    Set objChrome = New Selenium.ChromeDriver
    With objChrome
        .AddArgument "--headless" 'Khong hien thi Chrome khi chay thu tuc
        .Start 'Mo Chrome
        .Get strURL 'Truy cap trang web
        'Xac dinh cac hang du lieu
        Set objTRs = .FindElementById("main_table_countries_today").FindElementByTag("tbody").FindElementsByTag("tr")
    End With
    lngRowCount = objTRs.Count 'Dem so hang du lieu
    lngColumnCount = objTRs(1).FindElementsByTag("td").Count 'Dem so cot du lieu
    ReDim Preserve Data(1 To lngRowCount, 1 To lngColumnCount) 'Khai bao mang hai chieu de chua du lieu lay tu trang web
    r = 1 'Vi tri hang trong tap tin Excel
    'Duyet qua tung hang
    For Each objTR In objTRs
        c = 1
        Set objTDs = objTR.FindElementsByTag("td")
        'Duyet qua tung cot
        For Each objTD In objTDs
            Data(r, c) = objTD.Text 'Lay du lieu cua hang
            c = c + 1
        Next
        r = r + 1
    Next
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Dua du lieu lay tu trang web vao tap tin Excel
    Range(Cells(lngLastRow + 1, 1), Cells(lngLastRow + lngRowCount, lngColumnCount)).Value = Data
    Set objChrome = Nothing
    Set objTRs = Nothing
    Set objTR = Nothing
    Set objTDs = Nothing
    Set objTD = Nothing
End Sub
