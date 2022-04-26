'Class Module 'MicrosoftSQLServerConnection'
Option Explicit

Private m_User As String, m_Password As String, m_Server As String
Private m_Port As String
Private m_Database As String
Public Enum UseTrustedConnection
    DoNotUseTrustedConnection = 0
    UseTrustedConnection = 1
End Enum

Public Property Let User(value As String)
    m_User = value
End Property

Public Property Get User() As String
    User = m_User
End Property

Public Property Let Password(value As String)
    m_Password = value
End Property

Public Property Get Password() As String
    m_Password = Password
End Property

Public Property Let Server(value As String)
    m_Server = value
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Let Port(value As String)
    m_Port = value
End Property

Public Property Get Port() As String
    Port = m_Port
End Property

Public Property Let Database(value As String)
    m_Database = value
End Property

Public Property Get Database() As String
    Database = m_Database
End Property

Public Function RunSQL(SQLStatement As String, Optional TrustedConnection As UseTrustedConnection = 0, Optional User As Variant, Optional Password As Variant, Optional Server As Variant, Optional Port As Variant) As ADODB.Recordset
    Dim objCnn As ADODB.Connection
    Dim objCmd As ADODB.Command
    Dim strCnnStr As String
    If IsMissing(User) Then User = m_User
    If IsMissing(Password) Then Password = m_Password
    If IsMissing(Server) Then Server = m_Server
    If IsMissing(Port) Then Port = m_Port
    strCnnStr = "Driver={SQL Server};Server=" & Server & SpecifiedPort(Port) & ";Database=" & Database & ";User Id=" & User & ";Password=" & Password & ";"
    Select Case TrustedConnection
        Case UseTrustedConnection.UseTrustedConnection
            strCnnStr = strCnnStr & "Trusted_Connection = Yes;"
        Case Else
            strCnnStr = strCnnStr & "Trusted_Connection = No;"
    End Select
    Set objCnn = New ADODB.Connection
    With objCnn
        .ConnectionString = strCnnStr
        .ConnectionTimeout = 5
        .Open
    End With
    Set objCmd = New ADODB.Command
    With objCmd
        .ActiveConnection = objCnn
        .CommandText = SQLStatement
        Set RunSQL = .Execute
    End With
    Set objCnn = Nothing
    Set objCmd = Nothing
End Function

Private Function SpecifiedPort(PortNum As Variant) As String
    If PortNum = vbNullString Then
        SpecifiedPort = vbNullString
    Else: SpecifiedPort = "," & PortNum
    End If
End Function

