Option Explicit
'Khai báo biến toàn cục objChrome bên ngoài thủ tục sẽ ngăn không cho Chrome bị đóng lại sau khi kết thúc thủ tục
'Public Sub objChrome As Selenium.ChromeDriver As Selenium.ChromeDriver 

Public Sub GetTableDataFromWeb()
    'Khai báo biến objChrome bên trong thủ tục sẽ khiến cho Chrome bị đóng lại ngay sau khi kết thúc thủ tục
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
        .AddArgument "--headless" 'Đặt tham số này để Chrome không hiển thị khi chạy
        .Start 'Mở Chrome
        .Get strURL 'Truy cập trang web cần lấy dữ liệu
        'Xác định các hàng của bảng dữ liệu cần lấy
        Set objTRs = .FindElementById("main_table_countries_today").FindElementByTag("tbody").FindElementsByTag("tr")
    End With
    lngRowCount = objTRs.Count 'Đếm số hàng dữ liệu
    lngColumnCount = objTRs(1).FindElementsByTag("td").Count 'Đếm số cột dữ liệu
    'Khai báo mảng hai chiều để chứa dữ liệu lấy từ trang web
    ReDim Preserve Data(1 To lngRowCount, 1 To lngColumnCount) 
    r = 1 'Vi tri hang trong tap tin Excel
    'Duyệt qua từng hàng trong bảng
    For Each objTR In objTRs
        c = 1
        Set objTDs = objTR.FindElementsByTag("td")
        'Duyệt qua từng cột trong bảng
        For Each objTD In objTDs
            Data(r, c) = objTD.Text 'Lấy dữ liệu
            c = c + 1
        Next
        r = r + 1
    Next
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Đưa dữ liệu trong mảng hai chiều ra tập tin Excel
    Range(Cells(lngLastRow + 1, 1), Cells(lngLastRow + lngRowCount, lngColumnCount)).Value = Data
    Set objChrome = Nothing
    Set objTRs = Nothing
    Set objTR = Nothing
    Set objTDs = Nothing
    Set objTD = Nothing
End Sub
