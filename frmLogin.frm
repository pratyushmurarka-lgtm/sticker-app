VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QR Generator - Security Login"
   ClientHeight    =   1830
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2325
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Username:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call CloseCn
    End
End Sub

Private Sub cmdOK_Click()
    Dim rs As New ADODB.Recordset
    Dim userStr As String
    Dim passStr As String
    Dim sql As String
    
    userStr = Trim$(txtUsername.Text)
    passStr = Trim$(txtPassword.Text)
    
    If userStr = "" Or passStr = "" Then
        MsgBox "Please enter both username and password.", vbExclamation, "Missing Information"
        Exit Sub
    End If
    
    ' Connect to database and load roles
    On Error GoTo ErrLogin
    sql = "SELECT UserID, UserRole FROM UserMaster WHERE UserID='" & Replace(userStr, "'", "''") & _
          "' AND PasswordHash='" & Replace(passStr, "'", "''") & "'"
          
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        ' Successful Login
        GlobalUserID = rs!UserID
        GlobalUserRole = rs!UserRole
        rs.Close
        
        ' Load Permissions based on role
        If GlobalUserRole = "Admin" Then
            GlobalCanGenerateQR = True
            GlobalCanPrintNormal = True
            GlobalCanImportPDF = True
            GlobalCanReprint = True
            GlobalCanManageUsers = True
            GlobalCanViewReports = True
        ElseIf GlobalUserRole = "Supervisor" Then
            GlobalCanGenerateQR = True
            GlobalCanPrintNormal = True
            GlobalCanImportPDF = True
            GlobalCanReprint = True
            GlobalCanManageUsers = False
            GlobalCanViewReports = True
        Else ' Operator
            GlobalCanGenerateQR = True
            GlobalCanPrintNormal = True
            GlobalCanImportPDF = False
            GlobalCanReprint = False
            GlobalCanManageUsers = False
            GlobalCanViewReports = False
        End If
        
        ' Load Main form and close login
        frmMain.Show
        Unload Me
    Else
        rs.Close
        MsgBox "Invalid username or password. Please try again.", vbCritical, "Login Failed"
        txtPassword.Text = ""
        txtUsername.SetFocus
    End If
    Exit Sub
    
ErrLogin:
    MsgBox "Error during login validation: " & Err.Description, vbCritical, "Login Error"
End Sub

Private Sub Form_Load()
    ' Connect to database first
    Call OpenCn
End Sub
