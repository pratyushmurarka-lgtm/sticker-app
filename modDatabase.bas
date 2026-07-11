Attribute VB_Name = "modDatabase"
Option Explicit

' Global ADODB Connection Object
Public Cn As ADODB.Connection

' Global Logged-in User Session Details
Public GlobalUserID As String
Public GlobalUserRole As String
Public GlobalCanGenerateQR As Boolean
Public GlobalCanPrintNormal As Boolean
Public GlobalCanImportPDF As Boolean
Public GlobalCanReprint As Boolean
Public GlobalCanManageUsers As Boolean
Public GlobalCanViewReports As Boolean

' API Types and Declarations for Process Execution (ShellAndWait)
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, _
    ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    lpEnvironment As Any, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Const INFINITE As Long = &HFFFFFFFF
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_HIDE As Long = 0

' Sleep function API
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Open Connection to SQL Server 2012 Express
Public Sub OpenCn()
    On Error GoTo ErrHand
    
    Set Cn = New ADODB.Connection
    Cn.CursorLocation = adUseClient
    
    Dim connStr As String
    Dim serverName As String: serverName = "ERPINTECH\SQLEXPRESS"
    Dim dbName As String: dbName = "Es_Comp0001_2026"
    Dim uName As String: uName = "sa"
    Dim pWord As String: pWord = "12345678"
    
    Dim configPath As String
    configPath = "C:\dbconfig.txt"
    
    If FileExists(configPath) Then
        Dim fNo As Integer
        Dim line As String
        Dim pos As Long
        Dim key As String
        Dim val As String
        
        Dim dbMode As String: dbMode = "TEST"
        Dim liveServer As String: liveServer = ""
        Dim liveDB As String: liveDB = ""
        Dim liveUser As String: liveUser = ""
        Dim livePass As String: livePass = ""
        Dim testServer As String: testServer = ""
        Dim testDB As String: testDB = ""
        Dim testUser As String: testUser = ""
        Dim testPass As String: testPass = ""
        
        fNo = FreeFile
        Open configPath For Input As #fNo
        Do While Not EOF(fNo)
            Line Input #fNo, line
            line = Trim$(line)
            If Len(line) > 0 And InStr(line, "=") > 0 Then
                pos = InStr(line, "=")
                key = UCase$(Trim$(Left$(line, pos - 1)))
                val = Trim$(Mid$(line, pos + 1))
                
                Select Case key
                    Case "MODE"
                        dbMode = UCase$(val)
                    Case "LIVE_SERVER"
                        liveServer = val
                    Case "LIVE_DATABASE"
                        liveDB = val
                    Case "LIVE_USER"
                        liveUser = val
                    Case "LIVE_PASSWORD"
                        livePass = val
                    Case "TEST_SERVER"
                        testServer = val
                    Case "TEST_DATABASE"
                        testDB = val
                    Case "TEST_USER"
                        testUser = val
                    Case "TEST_PASSWORD"
                        testPass = val
                End Select
            End If
        Loop
        Close #fNo
        
        If dbMode = "LIVE" Then
            serverName = liveServer
            dbName = liveDB
            uName = liveUser
            pWord = livePass
        Else
            serverName = testServer
            dbName = testDB
            uName = testUser
            pWord = testPass
        End If
    End If
    
    connStr = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";User ID=" & uName & ";Password=" & pWord & ";"
    Cn.Open connStr
    Exit Sub
    
ErrHand:
    MsgBox "Failed to connect to Database: " & Err.Description, vbCritical, "Database Connection Error"
    End
End Sub

' Close Connection
Public Sub CloseCn()
    On Error Resume Next
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Set Cn = Nothing
End Sub

' Function to execute an external application and wait for it to complete
Public Function ShellAndWait(ByVal cmdLine As String, ByVal timeoutMs As Long) As Boolean
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim ret As Long
    
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    
    ' Execute CreateProcess
    ret = CreateProcess(vbNullString, cmdLine, 0, 0, 1, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
    
    If ret <> 0 Then
        ' Wait until process finishes or timeout occurs
        ret = WaitForSingleObject(pi.hProcess, timeoutMs)
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
        
        If ret = 0 Then
            ShellAndWait = True ' Success
        Else
            ShellAndWait = False ' Timeout or error
        End If
    Else
        ShellAndWait = False ' Failed to spawn process
    End If
End Function

' Utility to convert Millimeters to Twips
Public Function MMToTwips(ByVal mm As Single) As Single
    ' 1 mm = 56.7 twips
    MMToTwips = mm * 56.7
End Function

' Utility to check if file exists
Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function
