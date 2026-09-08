Option Explicit

Private m_strConnection As String
Private m_blnInitialized As Boolean

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Sub InitializeConnection(Optional ByVal pIniPath As String = "")
    
    Dim iniFile As String
    Dim server As String
    Dim database As String
    Dim user As String
    Dim password As String
    Dim winAuth As Boolean
    
    If pIniPath = "" Then
        iniFile = App.Path & "\Config.ini"
    Else
        iniFile = pIniPath
    End If
    
    If Dir(iniFile) = "" Then
        CreateDefaultIni iniFile
    End If
    
    server = ReadIniValue(iniFile, "Database", "Server", "localhost")
    database = ReadIniValue(iniFile, "Database", "Database", "FinancialTransactionDB")
    user = ReadIniValue(iniFile, "Database", "User", "sa")
    password = ReadIniValue(iniFile, "Database", "Password", "")
    winAuth = (ReadIniValue(iniFile, "Database", "WindowsAuthentication", "0") = "1")
    
    If winAuth Then
        m_strConnection = "Provider=SQLOLEDB.1;" & _
                          "Data Source=" & server & ";" & _
                          "Initial Catalog=" & database & ";" & _
                          "Integrated Security=SSPI;" & _
                          "Persist Security Info=False"
    Else
        m_strConnection = "Provider=SQLOLEDB.1;" & _
                          "Data Source=" & server & ";" & _
                          "Initial Catalog=" & database & ";" & _
                          "User ID=" & user & ";" & _
                          "Password=" & password & ";" & _
                          "Persist Security Info=False"
    End If
    
    m_blnInitialized = True
    
End Sub

Private Function ReadIniValue(ByVal pFile As String, _
                              ByVal pSection As String, _
                              ByVal pKey As String, _
                              ByVal pDefault As String) As String
    
    Dim buffer As String * 255
    Dim ret As Long
    
    ret = GetPrivateProfileString(pSection, pKey, pDefault, buffer, 255, pFile)
    ReadIniValue = Left(buffer, ret)
    
End Function

Private Sub CreateDefaultIni(ByVal pFile As String)
    
    Dim f As Integer
    
    f = FreeFile
    
    Open pFile For Output As #f
    Print #f, "[Database]"
    Print #f, "Server=localhost"
    Print #f, "Database=FinancialTransactionDB"
    Print #f, "User=sa"
    Print #f, "Password=YourStrong@Password123"
    Print #f, "WindowsAuthentication=0"
    Close #f
    
End Sub

Public Function GetConnection() As ADODB.Connection
On Error GoTo ErrorHandler
    
    Dim cn As ADODB.Connection
    
    If Not m_blnInitialized Then
        InitializeConnection
    End If
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = m_strConnection
    cn.ConnectionTimeout = 15
    cn.CommandTimeout = 30
    cn.Open
    
    Set GetConnection = cn
    Exit Function
    
ErrorHandler:
    SaveLogErr "GetConnection", Err.Description
    Set GetConnection = Nothing
End Function

Public Function TestConnection() As Boolean
On Error GoTo ErrorHandler
    
    Dim cn As ADODB.Connection
    
    Set cn = GetConnection
    If cn Is Nothing Then
        TestConnection = False
        Exit Function
    End If
    
    TestConnection = (cn.State = adStateOpen)
    
    cn.Close
    Set cn = Nothing
    Exit Function
    
ErrorHandler:
    SaveLogErr "TestConnection", Err.Description
    TestConnection = False
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

Public Function QueryExec(ByVal pSql As String, ParamArray pParams()) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Long
    
    Set cn = GetConnection
    If cn Is Nothing Then
        Set QueryExec = Nothing
        Exit Function
    End If
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandText = pSql
        .CommandType = adCmdText
        .CommandTimeout = 30
        
        For i = LBound(pParams) To UBound(pParams) Step 2
            If i + 1 <= UBound(pParams) Then
                .Parameters.Append .CreateParameter(pParams(i), pParams(i + 1), adParamInput)
            End If
        Next i
    End With
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    
    Set rs.ActiveConnection = Nothing
    cn.Close
    
    Set QueryExec = rs
    Set cmd = Nothing
    Set cn = Nothing
    Exit Function
    
ErrorHandler:
    SaveLogErr "QueryExec", Err.Description, "SQL: " & pSql
    Set QueryExec = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

Public Function CommandExec(ByVal pSql As String, ParamArray pParams()) As Long
On Error GoTo ErrorHandler
    
    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim i As Long
    
    Set cn = GetConnection
    If cn Is Nothing Then
        CommandExec = -1
        Exit Function
    End If
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandText = pSql
        .CommandType = adCmdText
        .CommandTimeout = 30
        
        For i = LBound(pParams) To UBound(pParams) Step 2
            If i + 1 <= UBound(pParams) Then
                .Parameters.Append .CreateParameter(pParams(i), pParams(i + 1), adParamInput)
            End If
        Next i
        
        .Execute RecordsAffected:=CommandExec
    End With
    
    cn.Close
    Set cmd = Nothing
    Set cn = Nothing
    Exit Function
    
ErrorHandler:
    SaveLogErr "CommandExec", Err.Description, "SQL: " & pSql
    CommandExec = -1
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function