VERSION 5.00

Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"

Begin VB.Form frmMain 

   Caption         =   "QR Code Generation & Print Management System - Dashboard"

   ClientHeight    =   9240

   ClientLeft      =   60

   ClientTop       =   345

   ClientWidth     =   14865

   LinkTopic       =   "Form1"

   ScaleHeight     =   9240

   ScaleWidth      =   14865

   StartUpPosition =   2  'CenterScreen

   WindowState     =   2  'Maximized

   Begin VB.PictureBox picSidebar 

      Align           =   3  'Align Left

      BackColor       =   &H002D2724&

      BorderStyle     =   0  'None

      Height          =   9240

      Left            =   0

      ScaleHeight     =   9240

      ScaleWidth      =   2800

      TabIndex        =   15

      Top             =   0

      Width           =   2800

      Begin VB.CommandButton cmdTabReports 

         Caption         =   "Operational Reports"

         Height          =   615

         Left            =   120

         TabIndex        =   26

         Top             =   3240

         Width           =   2535

      End

      Begin VB.CommandButton cmdTabUsers 

         Caption         =   "User Management"

         Height          =   615

         Left            =   120

         TabIndex        =   51

         Top             =   4080

         Width           =   2535

      End

      Begin VB.CommandButton cmdLogout 

         Caption         =   "Logout"

         Height          =   495

         Left            =   120

         TabIndex        =   21

         Top             =   8400

         Width           =   2535

      End

      Begin VB.CommandButton cmdTabCust 

         Caption         =   "Customer PDF Operations"

         Height          =   615

         Left            =   120

         TabIndex        =   17

         Top             =   2400

         Width           =   2535

      End

      Begin VB.CommandButton cmdTabGen 

         Caption         =   "Sequential Generation"

         Height          =   615

         Left            =   120

         TabIndex        =   16

         Top             =   1560

         Width           =   2535

      End

      Begin VB.Label lblTitle 

         Alignment       =   2  'Center

         BackColor       =   &H002D2724&

         Caption         =   "QR DASHBOARD"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   12

            Weight          =   700

         EndProperty

         ForeColor       =   &H00FFFFFF&

         Height          =   375

         Left            =   120

         TabIndex        =   23

         Top             =   240

         Width           =   2535

      End

      Begin VB.Label lblUserSession 

         Alignment       =   2  'Center

         BackColor       =   &H002D2724&

         Caption         =   "User: N/A"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   8.25

            Weight          =   400

         EndProperty

         ForeColor       =   &H00CCCCCC&

         Height          =   615

         Left            =   120

         TabIndex        =   20

         Top             =   720

         Width           =   2535

      End

   End

   Begin VB.Frame fraReports 

      Caption         =   "Operational Reports Dashboard"

      BeginProperty Font 

         Name            =   "Arial"

         Size            =   9.75

         Weight          =   700

      EndProperty

      Height          =   8895

      Left            =   3000

      TabIndex        =   27

      Top             =   120

      Width           =   12200

      Visible         =   0   'False

      Begin VB.ComboBox cboReportType 

         Height          =   315

         Left            =   1800

         Style           =   2  'Dropdown List

         TabIndex        =   28

         Top             =   420

         Width           =   3800

      End

      Begin VB.CommandButton cmdRunReport 

         Caption         =   "Run Report"

         Height          =   495

         Left            =   5800

         TabIndex        =   29

         Top             =   360

         Width           =   2000

      End

      Begin VB.CommandButton cmdExportCSV 

         Caption         =   "Export CSV"

         Height          =   495

         Left            =   8000

         TabIndex        =   30

         Top             =   360

         Width           =   2000

      End

      Begin MSFlexGridLib.MSFlexGrid fgReports 

         Height          =   7300

         Left            =   360

         TabIndex        =   31

         Top             =   1200

         Width           =   20300

         _ExtentX        =   35807

         _ExtentY        =   12876

         _Version        =   393216

         Cols            =   5

         FixedCols       =   0

         FocusRect       =   0

         SelectionMode   =   1

      End

      Begin VB.Label lblReportType 

         Caption         =   "Select Report:"

         Height          =   255

         Left            =   360

         TabIndex        =   32

         Top             =   480

         Width           =   1335

      End

   End

   Begin VB.Frame fraUsers 

      Caption         =   "User Management Control Panel"

      BeginProperty Font 

         Name            =   "Arial"

         Size            =   9.75

         Weight          =   700

      EndProperty

      Height          =   8895

      Left            =   3000

      TabIndex        =   52

      Top             =   120

      Width           =   12200

      Visible         =   0   'False

      Begin MSFlexGridLib.MSFlexGrid fgUsers 

         Height          =   5500

         Left            =   360

         TabIndex        =   53

         Top             =   1320

         Width           =   11400

         _ExtentX        =   20108

         _ExtentY        =   9702

         _Version        =   393216

         Cols            =   3

         FixedCols       =   0

         FocusRect       =   0

         SelectionMode   =   1

      End

      Begin VB.Label lblUserID 

         BackStyle       =   0  'Transparent

         Caption         =   "User ID:"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9.75

            Weight          =   700

         EndProperty

         Height          =   255

         Left            =   360

         TabIndex        =   54

         Top             =   7100

         Width           =   1000

      End

      Begin VB.TextBox txtUserID 

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9.75

            Weight          =   400

         EndProperty

         Height          =   350

         Left            =   1500

         TabIndex        =   55

         Top             =   7050

         Width           =   2500

      End

      Begin VB.Label lblPassword 

         BackStyle       =   0  'Transparent

         Caption         =   "Password:"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9.75

            Weight          =   700

         EndProperty

         Height          =   255

         Left            =   4300

         TabIndex        =   56

         Top             =   7100

         Width           =   1000

      End

      Begin VB.TextBox txtPassword 

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9.75

            Weight          =   400

         EndProperty

         Height          =   350

         Left            =   5400

         TabIndex        =   57

         Top             =   7050

         Width           =   2500

      End

      Begin VB.Label lblUserRole 

         BackStyle       =   0  'Transparent

         Caption         =   "Role:"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9.75

            Weight          =   700

         EndProperty

         Height          =   255

         Left            =   8200

         TabIndex        =   58

         Top             =   7100

         Width           =   800

      End

      Begin VB.ComboBox cboUserRole 

         Height          =   315

         Left            =   9100

         Style           =   2  'Dropdown List

         TabIndex        =   59

         Top             =   7050

         Width           =   2000

      End

      Begin VB.CommandButton cmdAddUser 

         Caption         =   "Add User"

         Height          =   495

         Left            =   360

         TabIndex        =   60

         Top             =   7700

         Width           =   1800

      End

      Begin VB.CommandButton cmdUpdateUser 

         Caption         =   "Update User"

         Height          =   495

         Left            =   2400

         TabIndex        =   61

         Top             =   7700

         Width           =   1800

      End

      Begin VB.CommandButton cmdDeleteUser 

         Caption         =   "Delete User"

         Height          =   495

         Left            =   4440

         TabIndex        =   62

         Top             =   7700

         Width           =   1800

      End

      Begin VB.CommandButton cmdClearFields 

         Caption         =   "Clear Fields"

         Height          =   495

         Left            =   6480

         TabIndex        =   63

         Top             =   7700

         Width           =   1800

      End

   End

   Begin VB.Frame fraCust 

      Caption         =   "Customer-Provided PDF QR Code Management"

      BeginProperty Font 

         Name            =   "Arial"

         Size            =   9.75

         Weight          =   700

      EndProperty

      Height          =   8895

      Left            =   3000

      TabIndex        =   11

      Top             =   120

      Width           =   12200

      Visible         =   0   'False

      Begin VB.PictureBox picProgressOuter 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12200
         ScaleHeight     =   270
         ScaleWidth      =   4970
         TabIndex        =   45
         Top             =   420
         Visible         =   0  'False
         Width           =   5000
         Begin VB.PictureBox picProgressInner 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            Height          =   260
            Left            =   0
            ScaleHeight     =   260
            ScaleWidth      =   0
            TabIndex        =   46
            Top             =   0
            Width           =   0
         End
      End
      Begin VB.Label lblProgressText 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Weight          =   700
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   12200
         TabIndex        =   47
         Top             =   800
         Visible         =   0  'False
         Width           =   6000
      End
      Begin VB.Label lblRegex 
         BackStyle       =   0  'Transparent
         Caption         =   "QR Code Format (Regex):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Weight          =   700
         EndProperty
         Height          =   255
         Left            =   5900
         TabIndex        =   48
         Top             =   580
         Width           =   2400
      End
      Begin VB.TextBox txtRegex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Weight          =   400
         EndProperty
         Height          =   350
         Left            =   8400
         TabIndex        =   49
         Top             =   530
         Width           =   3500
         Text            =   "\b\d{16}\b"
      End
      Begin VB.Label lblRegexHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "Default formats: \b\d{16}\b (16-digits), \bHD\d{7}\b (HD codes), \bCE\d{7}\b|\bU-\d{5}\b (CE codes)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Weight          =   400
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   5900
         TabIndex        =   50
         Top             =   950
         Width           =   9000
      End
      Begin VB.CommandButton cmdImportPDF 

         Caption         =   "Import Customer PDF"

         Height          =   550

         Left            =   360

         TabIndex        =   12

         Top             =   480

         Width           =   2500

      End

      Begin VB.CommandButton cmdPrintPDF 

         Caption         =   "Print Selected PDF"

         Height          =   550

         Left            =   3120

         TabIndex        =   13

         Top             =   480

         Width           =   2500

      End

      Begin MSFlexGridLib.MSFlexGrid fgPDF 

         Height          =   7200

         Left            =   360

         TabIndex        =   14

         Top             =   1400

         Width           =   20300

         _ExtentX        =   35807

         _ExtentY        =   12700

         _Version        =   393216

         Cols            =   7

         FixedCols       =   0

         FocusRect       =   0

         SelectionMode   =   1

      End

   End

   Begin VB.Frame fraGen 

      Caption         =   "Sequential QR Code Generation & Printing"

      BeginProperty Font 

         Name            =   "Arial"

         Size            =   9.75

         Weight          =   700

      EndProperty

      Height          =   8895

      Left            =   3000

      TabIndex        =   0

      Top             =   120

      Width           =   21000

      Begin VB.TextBox txtSearch 

         Height          =   315

         Left            =   1800

         TabIndex        =   1

         Text            =   ""

         Top             =   420

         Width           =   1695

      End

      Begin VB.ComboBox cboItem 

         Height          =   315

         Left            =   5800

         TabIndex        =   2

         Top             =   420

         Width           =   5500

      End

      Begin VB.TextBox txtQty 

         Height          =   375

         Left            =   1800

         TabIndex        =   3

         Text            =   "1"

         Top             =   1220

         Width           =   1200

      End

      Begin VB.CommandButton cmdGenerate 

         Caption         =   "Generate QR Codes"

         Height          =   495

         Left            =   3200

         TabIndex        =   4

         Top             =   1220

         Width           =   2200

      End

      Begin VB.CommandButton cmdExportPDF 

         Caption         =   "Print Current Batch"

         Height          =   495

         Left            =   5600

         TabIndex        =   5

         Top             =   1220

         Width           =   2500

      End

      Begin VB.PictureBox piclayout 

         AutoRedraw      =   -1  'True

         BackColor       =   &H00FFFFFF&

         Height          =   750

         Left            =   17600

         ScaleMode       =   1

         TabIndex        =   6

         Top             =   1100

         Width           =   3000

      End

      Begin VB.CommandButton cmdLoadReprint 

         Caption         =   "Load Reprint List"

         Height          =   495

         Left            =   360

         TabIndex        =   7

         Top             =   1800

         Width           =   2200

      End

      Begin VB.CommandButton cmdReprint 

         Caption         =   "Reprint Selected"

         Height          =   495

         Left            =   2760

         TabIndex        =   8

         Top             =   1800

         Width           =   2200

      End





      Begin MSFlexGridLib.MSFlexGrid fg 

         Height          =   6000

         Left            =   360

         TabIndex        =   9

         Top             =   2520

         Width           =   20300

         _ExtentX        =   35807

         _ExtentY        =   10583

         _Version        =   393216

         Cols            =   5

         FixedCols       =   0

         FocusRect       =   0

         SelectionMode   =   1

      End

      Begin VB.Label lblSearch 

         Caption         =   "Search Model:"

         Height          =   255

         Left            =   360

         TabIndex        =   24

         Top             =   480

         Width           =   1335

      End

      Begin VB.Label lblModel 

         Caption         =   "Select Model:"

         Height          =   255

         Left            =   4500

         TabIndex        =   10

         Top             =   480

         Width           =   1200

      End

      Begin VB.Label lblLastSerial 

         Caption         =   "Last Generated QR: N/A"

         BeginProperty Font 

            Name            =   "Arial"

            Size            =   9

            Weight          =   700

         EndProperty

         ForeColor       =   &H00C00000&

         Height          =   255

         Left            =   360

         TabIndex        =   39

         Top             =   820

         Width           =   20300

      End

      Begin VB.Label lblQty 

         Caption         =   "Quantity:"

         Height          =   255

         Left            =   360

         TabIndex        =   25

         Top             =   1280

         Width           =   1200

      End

   End

End

Attribute VB_Name = "frmMain"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Option Explicit



' Collection to hold generated QR values for the current run

Private colQRList As Collection



' Internal list mapping for Combobox autocomplete

Dim FullItemList() As String

Dim ItemCount As Long



' Windows API to print files directly

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long



Private Const SW_HIDE As Long = 0

' Native Win32 API declarations for asynchronous process polling
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const STILL_ACTIVE As Long = &H103

' Module-level state variables for print dialog page range selection
Private m_fromPage As Long
Private m_toPage As Long
Private m_isRangeSelected As Boolean



' Path to zint utility

Private Const ZINT_PATH As String = "C:\QR_Test\zint.exe"



' Native Windows Open File Dialog API & Type Definitions

Private Type OPENFILENAME

    lStructSize As Long

    hwndOwner As Long

    hInstance As Long

    lpstrFilter As String

    lpstrCustomFilter As String

    nMaxCustFilter As Long

    nFilterIndex As Long

    lpstrFile As String

    nMaxFile As Long

    lpstrFileTitle As String

    nMaxFileTitle As Long

    lpstrInitialDir As String

    lpstrTitle As String

    flags As Long

    nFileOffset As Integer

    nFileExtension As Integer

    lpstrDefExt As String

    lCustData As Long

    lpfnHook As Long

    lpTemplateName As String

End Type



Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
    pOpenfilename As OPENFILENAME) As Long



Private Const OFN_FILEMUSTEXIST As Long = &H1000

Private Const OFN_PATHMUSTEXIST As Long = &H800



' API to set System Default Printer

Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" ( _
    ByVal pszPrinter As String) As Long



' Windows API and Types for Print Dialog Box

Private Type PRINTDLG_TYPE

    lStructSize As Long

    hwndOwner As Long

    hDevMode As Long

    hDevNames As Long

    hdc As Long

    flags As Long

    nFromPage As Integer

    nToPage As Integer

    nMinPage As Integer

    nMaxPage As Integer

    nCopies As Integer

    hInstance As Long

    lCustData As Long

    lpfnPrintHook As Long

    lpfnSetupHook As Long

    lpPrintTemplateName As String

    lpSetupTemplateName As String

    hPrintTemplate As Long

    hSetupTemplate As Long

End Type



Private Type DEVNAMES_TYPE

    wDriverOffset As Integer

    wDeviceOffset As Integer

    wOutputOffset As Integer

    wDefault As Integer

End Type



Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" ( _
    pPrintdlg As PRINTDLG_TYPE) As Long



Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)



' Windows API and Constant for ComboBox dropdown handling

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long



Private Const CB_SHOWDROPDOWN As Long = &H14F



' Form load routine

Private Sub Form_Load()

    ' Prepare layout preview canvas size in twips (100 x 25 mm)

    piclayout.AutoRedraw = True

    piclayout.ScaleMode = vbTwips

    piclayout.Width = MMToTwips(100)

    piclayout.Height = MMToTwips(25)

    

    Set colQRList = New Collection

    

    ' Populate Combobox and Grid

    Call LoadItemMaster

    Call LoadCustomerPDFGrid

    



    

    ' Populate Report Types

    cboReportType.Clear

    cboReportType.AddItem "1. Supervisor Reprint Audit Trail"

    cboReportType.AddItem "2. Registry vs Packing Scan Reconciliation"

    cboReportType.AddItem "3. Customer PDF Fulfillment Status"

    cboReportType.AddItem "4. Daily Production Summary"

    cboReportType.ListIndex = 0

    

    ' Set up user session label

    lblUserSession.Caption = "User: " & GlobalUserID & vbCrLf & "Role: " & GlobalUserRole & vbCrLf & "Items Loaded: " & ItemCount

    

    ' Enforce Role-based controls on UI elements

    cmdImportPDF.Enabled = GlobalCanImportPDF

    cmdTabReports.Enabled = GlobalCanViewReports

    cmdTabUsers.Visible = GlobalCanManageUsers

    

    ' Populate User Role dropdown

    cboUserRole.Clear

    cboUserRole.AddItem "Operator"

    cboUserRole.AddItem "Supervisor"

    cboUserRole.AddItem "Admin"

    

    ' Adjust Grid layouts

    Call InitializeGrids

    

    ' Activate Default Tab

    Call cmdTabGen_Click

    

    ' Update last printed serial initially

    Call UpdateLastPrintedSerial

    

    ' Force initial layout fit calculation

    Call Form_Resize

End Sub



' Form Resize Event Handler: Dynamically resizes frames, grids, and preview boxes based on user monitor resolution

Private Sub Form_Resize()

    On Error Resume Next

    Dim availWidth As Long

    Dim availHeight As Long

    

    availWidth = Me.ScaleWidth - picSidebar.Width - 450

    availHeight = Me.ScaleHeight - 300

    

    If availWidth < 5000 Then Exit Sub

    

    ' Resize Active Frame panels to fill screen width and height

    fraGen.Width = availWidth

    fraGen.Height = availHeight

    fraCust.Width = availWidth

    fraCust.Height = availHeight

    fraReports.Width = availWidth

    fraReports.Height = availHeight

    fraUsers.Width = availWidth

    fraUsers.Height = availHeight

    

    ' Resize FlexGrid controls proportionately to fill frames

    fg.Width = availWidth - 720

    fg.Height = availHeight - 3020

    

    fgPDF.Width = availWidth - 720

    fgPDF.Height = availHeight - 1760

    

    fgReports.Width = availWidth - 720

    fgReports.Height = availHeight - 1680

    

    fgUsers.Width = availWidth - 720

    fgUsers.Height = availHeight - 3300

    

    ' Position User Management controls and buttons dynamically

    Dim userControlsTop As Long

    userControlsTop = fgUsers.Top + fgUsers.Height + 150

    

    lblUserID.Top = userControlsTop + 50

    txtUserID.Top = userControlsTop

    lblPassword.Top = userControlsTop + 50

    txtPassword.Top = userControlsTop

    lblUserRole.Top = userControlsTop + 50

    cboUserRole.Top = userControlsTop

    

    Dim userButtonsTop As Long

    userButtonsTop = fgUsers.Top + fgUsers.Height + 700

    

    cmdAddUser.Top = userButtonsTop

    cmdUpdateUser.Top = userButtonsTop

    cmdDeleteUser.Top = userButtonsTop

    cmdClearFields.Top = userButtonsTop

    

    ' Position layout preview box dynamically in top-right of Generation frame

    piclayout.Left = availWidth - piclayout.Width - 360

    

    ' Stretch the Last Generated QR details label to full width

    lblLastSerial.Width = availWidth - 720

    

    ' Re-calculate and fit column widths proportionally

    Call FitGridColumns

End Sub



' ComboBox Selection Click Event

Private Sub cboItem_Click()

    Call UpdateLastPrintedSerial

End Sub



' Sidebar Navigation Controls

Private Sub cmdTabGen_Click()

    fraGen.Visible = True

    fraCust.Visible = False

    fraReports.Visible = False

    fraUsers.Visible = False

    cmdTabGen.FontBold = True

    cmdTabCust.FontBold = False

    cmdTabReports.FontBold = False

    cmdTabUsers.FontBold = False

End Sub



Private Sub cmdTabCust_Click()

    fraGen.Visible = False

    fraCust.Visible = True

    fraReports.Visible = False

    fraUsers.Visible = False

    cmdTabGen.FontBold = False

    cmdTabCust.FontBold = True

    cmdTabReports.FontBold = False

    cmdTabUsers.FontBold = False

End Sub



Private Sub cmdTabReports_Click()

    fraGen.Visible = False

    fraCust.Visible = False

    fraReports.Visible = True

    fraUsers.Visible = False

    cmdTabGen.FontBold = False

    cmdTabCust.FontBold = False

    cmdTabReports.FontBold = True

    cmdTabUsers.FontBold = False

    

    ' Initialize reports grid layout

    fgReports.Clear

    fgReports.Rows = 1

    fgReports.Cols = 5

End Sub



Private Sub cmdTabUsers_Click()

    fraGen.Visible = False

    fraCust.Visible = False

    fraReports.Visible = False

    fraUsers.Visible = True

    cmdTabGen.FontBold = False

    cmdTabCust.FontBold = False

    cmdTabReports.FontBold = False

    cmdTabUsers.FontBold = True

    

    Call ClearUserFields

    Call LoadUserGrid

End Sub



' Set up headers for FlexGrid controls

Private Sub InitializeGrids()

    ' Enable user-resizable columns in all grids

    fg.AllowUserResizing = 1 ' flexResizeColumns

    fgPDF.AllowUserResizing = 1 ' flexResizeColumns

    fgReports.AllowUserResizing = 1 ' flexResizeColumns

    

    ' Ensure both horizontal and vertical scrollbars are enabled

    fg.ScrollBars = 3 ' flexScrollBarBoth

    fgPDF.ScrollBars = 3 ' flexScrollBarBoth

    fgReports.ScrollBars = 3 ' flexScrollBarBoth



    ' Normal Reprint Grid (Expanded for Widescreen Layout)

    fg.Cols = 5

    

    fg.TextMatrix(0, 0) = "Row"

    fg.TextMatrix(0, 1) = "Model"

    fg.TextMatrix(0, 2) = "QR Details"

    fg.TextMatrix(0, 3) = "Created On"

    

    ' Apply left alignments to Model, QR Details, Created On, and Select Checkbox columns

    fg.ColAlignment(0) = 4 ' Row (Centered)

    fg.ColAlignment(1) = 1 ' Model (Left-aligned)

    fg.ColAlignment(2) = 1 ' QR Details (Left-aligned)

    fg.ColAlignment(3) = 1 ' Created On (Left-aligned)

    fg.ColAlignment(4) = 1 ' Select Checkbox (Left-aligned)

    

    ' Set select column header as Wingdings empty checkbox (o)

    fg.Row = 0

    fg.Col = 4

    fg.CellFontName = "Wingdings"

    fg.CellFontSize = 11

    fg.CellAlignment = 1 ' Left-aligned to match data checkboxes

    fg.TextMatrix(0, 4) = Chr$(111)

    

    ' PDF Grid (Expanded for Widescreen Layout)

    fgPDF.Cols = 7

    fgPDF.ColWidth(0) = 800

    fgPDF.ColWidth(1) = 3000

    fgPDF.ColWidth(2) = 11000

    fgPDF.ColWidth(3) = 1500

    fgPDF.ColWidth(4) = 1500

    fgPDF.ColWidth(5) = 1000

    fgPDF.ColWidth(6) = 1200

    

    fgPDF.TextMatrix(0, 0) = "ID"

    fgPDF.TextMatrix(0, 1) = "Model"

    fgPDF.TextMatrix(0, 2) = "File Name"

    fgPDF.TextMatrix(0, 3) = "Start Sr"

    fgPDF.TextMatrix(0, 4) = "End Sr"

    fgPDF.TextMatrix(0, 5) = "Qty"

    fgPDF.TextMatrix(0, 6) = "Printed?"

End Sub



' Load items into cboItem from ItemMaster

Private Sub LoadItemMaster()

    Dim rs As New ADODB.Recordset

    Dim sql As String

    Dim i As Long

    

    On Error GoTo ErrLoadMaster

    ' Query using exact column casing matching the database schema

    sql = "SELECT CodeStr, Name FROM ItemMaster WHERE Blocked = 0 AND ItemType IN (81,13414,94,4796,4846,82,22) ORDER BY CodeStr"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly

    

    ItemCount = rs.RecordCount

    If ItemCount > 0 Then

        ReDim FullItemList(0 To ItemCount - 1)

        cboItem.Clear

        i = 0

        Do While Not rs.EOF

            FullItemList(i) = rs!CodeStr & " - " & rs!Name

            cboItem.AddItem FullItemList(i)

            rs.MoveNext

            i = i + 1

        Loop

    Else

        ItemCount = 0

        cboItem.Clear

    End If

    rs.Close

    

    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0

    Exit Sub

    

ErrLoadMaster:

    ' Fallback items for local compilation verification

    ItemCount = 3

    ReDim FullItemList(0 To 2)

    FullItemList(0) = "44871WBD10AX26N - SWH WARMBOY WKT 10L"

    FullItemList(1) = "CE-000000505 - CE UNIT MAIN"

    FullItemList(2) = "WP01-TPL10 - TEST MODEL B"

    

    cboItem.Clear

    For i = 0 To 2

        cboItem.AddItem FullItemList(i)

    Next i

    cboItem.ListIndex = 0

End Sub



' --- SYSTEM RIGHTS VALIDATION ROUTINE ---

Private Function IsValidForPrinting(qrValue As String, ByRef errorMessage As String) As Boolean

    Dim rs As New ADODB.Recordset

    Dim sql As String

    

    ' 1. Check In_QRmaster (registry check)

    sql = "SELECT QRValue FROM In_QRmaster WHERE QRValue = '" & Replace(qrValue, "'", "''") & "'"

    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then

        rs.Close

        errorMessage = "QR code already exists in QR registry (In_QRmaster)."

        IsValidForPrinting = False

        Exit Function

    End If

    rs.Close

    

    ' 2. Check esjobscandata (Scan audit check - STRICTLY READ-ONLY SELECT QUERY)

    sql = "SELECT COUNT(1) AS ScanCount FROM esjobscandata WHERE ItemDetail = '" & Replace(qrValue, "'", "''") & "'"

    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    If rs!ScanCount > 0 Then

        rs.Close

        errorMessage = "QR code has already been scanned on a product box (esjobscandata)."

        IsValidForPrinting = False

        Exit Function

    End If

    rs.Close

    

    IsValidForPrinting = True

End Function



' Retrieve Next Serial Number from database

Private Function GetNextSerialNumber(itemType As Long, itemCode As String, itemBrand As Long) As Long

    Dim rs As New ADODB.Recordset

    Dim sql As String

    Dim qrMonth As String

    Dim NextSerial As Long

    

    qrMonth = Format(Date, "mmyyyy")

    

    ' Special Brand 1612 / Type 81 shared sequence condition

    If itemBrand = 1612 And itemType = 81 Then

        sql = "SELECT ISNULL(MAX(q.SerialNo), 0) + 1 AS NextSerial " & _
              "FROM In_QRmaster q " & _
              "INNER JOIN ItemMaster i ON q.ItemCode = i.CodeStr " & _
              "WHERE i.ItemBrand = 1612 AND i.ItemType = 81 AND q.QRMonth = '" & qrMonth & "'"

    Else

        ' Standard serialization grouping rules

        If itemType = 97 Or itemType = 4797 Or itemType = 10991 Then

            sql = "SELECT ISNULL(MAX(SerialNo), 0) + 1 AS NextSerial FROM In_QRmaster WHERE ItemCode='" & Replace(itemCode, "'", "''") & "'"

        Else

            sql = "SELECT ISNULL(MAX(SerialNo), 0) + 1 AS NextSerial FROM In_QRmaster WHERE ItemCode='" & Replace(itemCode, "'", "''") & "' AND QRMonth='" & qrMonth & "'"

        End If

    End If

    

    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    NextSerial = rs!NextSerial

    rs.Close

    

    GetNextSerialNumber = NextSerial

End Function



' Create candidate QR code values and run validators

Private Sub cmdGenerate_Click()

    Dim totalQty As Long

    Dim itemText As String

    Dim itemCode As String

    Dim i As Long

    Dim candidateQR As String

    Dim nextSerial As Long

    Dim errMsg As String

    Dim rsItem As New ADODB.Recordset

    Dim sql As String

    

    If Not GlobalCanGenerateQR Then

        MsgBox "Your account role does not have rights to generate QR codes.", vbCritical, "Access Denied"

        Exit Sub

    End If

    

    itemText = Trim$(cboItem.Text)

    If InStr(itemText, " - ") > 0 Then

        itemCode = Left$(itemText, InStr(itemText, " - ") - 1)

    Else

        itemCode = itemText

    End If

    

    totalQty = Val(txtQty.Text)

    If totalQty <= 0 Then

        MsgBox "Enter valid quantity.", vbExclamation

        Exit Sub

    End If

    

    ' Fetch full details from ItemMaster to support the 7-segment QR format

    sql = "SELECT Code, OtherDes, CodeStr, MRP, Name, ItemType, ItemBrand FROM ItemMaster WHERE CodeStr='" & Replace(itemCode, "'", "''") & "'"

    rsItem.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    

    If rsItem.EOF Then

        rsItem.Close

        MsgBox "Could not load item details from database.", vbCritical

        Exit Sub

    End If

    

    Dim itemDB_Code As Long

    Dim itemDB_OtherDes As String

    Dim itemDB_CodeStr As String

    Dim itemDB_MRP As Double

    Dim itemDB_Name As String

    Dim itemDB_Type As Long

    Dim itemDB_Brand As Long

    

    itemDB_Code = rsItem!Code

    itemDB_OtherDes = Trim$(IIf(IsNull(rsItem!OtherDes), "", rsItem!OtherDes))

    itemDB_CodeStr = rsItem!CodeStr

    itemDB_MRP = Val(IIf(IsNull(rsItem!MRP), 0, rsItem!MRP))

    itemDB_Name = rsItem!Name

    itemDB_Type = rsItem!ItemType

    itemDB_Brand = Val(IIf(IsNull(rsItem!ItemBrand), 0, rsItem!ItemBrand))

    rsItem.Close

    

    Set colQRList = New Collection

    Screen.MousePointer = vbHourglass

    

    For i = 1 To totalQty

        ' Request next serial number based on datatype itemtype and brand rules

        nextSerial = GetNextSerialNumber(itemDB_Type, itemDB_CodeStr, itemDB_Brand)

        

        ' Format: CODE|OtherDes|codestr|YYYYMM|MRP|NAME|SRNO

        candidateQR = itemDB_Code & "|" & itemDB_OtherDes & "|" & itemDB_CodeStr & "|" & _
                      Format(Now, "yyyymm") & "|" & itemDB_MRP & "|" & itemDB_Name & "|" & nextSerial

        

        ' Run validation

        If IsValidForPrinting(candidateQR, errMsg) Then

            ' Insert with Printed=0

            If InsertQRToDB(itemDB_CodeStr, candidateQR, nextSerial) Then

                colQRList.Add candidateQR

            Else

                ' Decrement to retry if insert failed

                i = i - 1

            End If

        Else

            ' Collision found - skip serial and notify

            MsgBox "Collision bypass triggered: " & errMsg, vbInformation

            i = i - 1

        End If

        DoEvents

    Next i

    

    Screen.MousePointer = vbDefault

    MsgBox "Generated " & colQRList.count & " serial numbers.", vbInformation

    DrawPagePreview 1

End Sub



' Helper database insert

Private Function InsertQRToDB(itemCode As String, qrValue As String, serialNo As Long) As Boolean

    On Error GoTo ErrIns

    Dim sql As String

    Dim qrMonth As String

    qrMonth = Format(Date, "mmyyyy")

    

    sql = "INSERT INTO In_QRmaster (ItemCode, QRMonth, SerialNo, QRValue, GeneratedOn, Printed, IsCustomerQR) VALUES (" & _
          "'" & Replace(itemCode, "'", "''") & "', '" & qrMonth & "', " & serialNo & ", '" & Replace(qrValue, "'", "''") & "', GETDATE(), 0, 0)"

    Cn.Execute sql

    InsertQRToDB = True

    Exit Function

ErrIns:

    InsertQRToDB = False

End Function



' Preview drawing

Private Sub DrawPagePreview(pageNumber As Long)

    piclayout.Cls

    If colQRList Is Nothing Then Exit Sub

    If colQRList.count = 0 Then Exit Sub

    

    ' Draw simple representation for layout previewing

    piclayout.CurrentX = 100

    piclayout.CurrentY = 100

    piclayout.Print "Previewing First QR of batch:"

    piclayout.CurrentX = 100

    piclayout.CurrentY = 400

    piclayout.Print colQRList(1)

End Sub



' Print/export current batch to PDF using CutePDF Writer (4 QR per row)

Private Sub cmdExportPDF_Click()
    Dim selectedPrinter As String

    If colQRList Is Nothing Or colQRList.count = 0 Then

        MsgBox "No QR codes generated to export. Click Generate first.", vbExclamation

        Exit Sub

    End If



    Dim prn As Printer

    Dim found As Boolean

    Dim total As Long, pages As Long, p As Long

    Dim i As Long, index As Long

    Dim labelSizeTw As Long, gapTw As Long, leftTw As Long, topTw As Long

    Dim qrpic As StdPicture

    Dim parts As Variant

    Dim qrvalue As String

    Dim p1 As String, p2 As String

    Dim txt As String

    

    ' Display standard Print Dialog Box


    selectedPrinter = DisplayPrintDlg()

    If selectedPrinter = "" Then Exit Sub ' Operator cancelled the dialog

    

    For Each prn In Printers

        If prn.DeviceName = selectedPrinter Then

            Set Printer = prn

            found = True

            Exit For

        End If

    Next



    If Not found Then

        MsgBox "The selected printer is offline or not found.", vbCritical

        Exit Sub

    End If



    ' Prepare label sizes in twips

    labelSizeTw = MMToTwips(16) ' 16 mm labels

    gapTw = MMToTwips(8)        ' 8 mm gaps

    topTw = MMToTwips(2)        ' 2 mm top margin



    Printer.ScaleMode = vbTwips

    piclayout.ScaleMode = vbTwips



    total = colQRList.count

    pages = (total + 3) \ 4



    ' Open print job

    For p = 1 To pages

        Printer.Font.Name = "Calibri"

        Printer.Font.Size = 8

        Printer.FontBold = False



        For i = 0 To 3

            index = (p - 1) * 4 + i + 1

            If index <= total Then

                '--- Generate QR bitmap image via zint.exe

                Set qrpic = GenerateQRPicture(colQRList(index))



                leftTw = MMToTwips(3) + i * (labelSizeTw + gapTw)



                '--- Print QR directly onto page

                If Not qrpic Is Nothing Then

                    Printer.PaintPicture qrpic, leftTw, topTw, labelSizeTw, labelSizeTw

                End If



                '--- Extract and format text label matching new 7-segment QR structure (e.g. CodeStr-SerialNo)

                qrvalue = colQRList(index)

                parts = Split(qrvalue, "|")

                If UBound(parts) >= 6 Then

                    p1 = Trim(parts(2)) ' CodeStr (part index 2)

                    p2 = Trim(parts(6)) ' SerialNo (part index 6)

                    txt = p1 & "-" & p2

                Else

                    txt = qrvalue

                End If



                '--- Print text description directly below QR code

                Printer.CurrentX = leftTw

                Printer.CurrentY = topTw + labelSizeTw + MMToTwips(1)



                ' Shrink text dynamically to fit within label width

                Do While Printer.TextWidth(txt) > labelSizeTw And Printer.Font.Size > 5

                    Printer.Font.Size = Printer.Font.Size - 0.5

                Loop



                Printer.Print txt

            End If

        Next i



        If p < pages Then Printer.NewPage

    Next p



    ' End document and send to CutePDF writer

    Printer.EndDoc



    MsgBox "Print job sent successfully to: " & Printer.DeviceName, vbInformation



    ' Mark generated QR entries as printed in database

    Cn.BeginTrans

    For i = 1 To colQRList.count

        Cn.Execute "UPDATE In_QRmaster SET Printed = 1 WHERE QRValue = '" & Replace(colQRList(i), "'", "''") & "'"

    Next i

    Cn.CommitTrans

    

    Set colQRList = New Collection

    

    ' Refresh last printed serial display

    Call UpdateLastPrintedSerial

End Sub



' --- CUSTOMER PDF IMPORT AND VALIDATION ROUTINES ---

Private Sub cmdImportPDF_Click()

    Dim pdfPath As String

    Dim tempFile As String

    Dim pythonCmd As String

    Dim success As Long

    Dim itemCode As String

    Dim regexPattern As String

    Dim fso As Object

    Dim ts As Object

    Dim qrValue As String

    Dim errMsg As String

    Dim colImportQRs As New Collection

    Dim duplicateFound As Boolean

    

    tempFile = App.Path & "\temp_extracted_qrs.txt"
    If Not GlobalCanImportPDF Then

        MsgBox "Access Denied: Your account role does not have rights to import PDFs.", vbCritical, "Access Denied"

        Exit Sub

    End If

    

    ' 1. Select PDF File using native Windows File Open Dialog box

    pdfPath = ShowOpenFileDialog()

    If pdfPath = "" Then Exit Sub ' Operator cancelled the dialog

    If Dir(pdfPath) = "" Then

        MsgBox "PDF file not found.", vbCritical

        Exit Sub

    End If

    

    ' 2. Select Model / Item Code

    Dim itemText As String

    itemText = Trim$(cboItem.Text)

    If InStr(itemText, " - ") > 0 Then

        itemCode = Left$(itemText, InStr(itemText, " - ") - 1)

    Else

        itemCode = itemText

    End If

    

    ' 3. Select / suggest regex pattern based on filename
    If InStr(LCase(pdfPath), "ce-") > 0 Then
        txtRegex.Text = "\bCE\d{7}\b|\bU-\d{5}\b"
    ElseIf InStr(LCase(pdfPath), "hd-") > 0 Then
        txtRegex.Text = "\bHD\d{7}\b"
    End If
    
    regexPattern = Trim$(txtRegex.Text)
    If regexPattern = "" Then
        MsgBox "Please enter a valid QR Code Regex Pattern.", vbCritical, "Format Error"
        Exit Sub
    End If
    If Dir(tempFile) <> "" Then Kill tempFile

    

    ' Clean up old log files
    On Error Resume Next
    If Dir(App.Path & "\parser_log.txt") <> "" Then Kill App.Path & "\parser_log.txt"
    On Error GoTo ErrDB
    
    ' Show progress UI components
    picProgressOuter.Visible = True
    lblProgressText.Visible = True
    picProgressInner.Width = 0
    lblProgressText.Caption = "Starting PDF text extraction..."
    DoEvents
    
    Dim progFile As String
    progFile = tempFile & ".prog"
    If Dir(progFile) <> "" Then Kill progFile
    
    ' Call compiled parser silently using WScript.Shell in non-blocking background mode
    pythonCmd = App.Path & "\extract_qrs.exe " & Chr$(34) & pdfPath & Chr$(34) & " " & Chr$(34) & tempFile & Chr$(34) & " " & Chr$(34) & regexPattern & Chr$(34)
    ' Call compiled parser silently using native VB6 Shell function (prevents WScript error 53)
    pythonCmd = Chr$(34) & App.Path & "\extract_qrs.exe" & Chr$(34) & " " & Chr$(34) & pdfPath & Chr$(34) & " " & Chr$(34) & tempFile & Chr$(34) & " " & Chr$(34) & regexPattern & Chr$(34)
    
    Dim pid As Long
    Dim hProcess As Long
    Dim exitCode As Long
    Dim startTime As Double
    startTime = Timer
    
    Screen.MousePointer = vbHourglass
    pid = Shell(pythonCmd, vbHide)
    
    If pid <> 0 Then
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
        If hProcess <> 0 Then
            Do
                Call GetExitCodeProcess(hProcess, exitCode)
                DoEvents
                Sleep 50 ' 50ms pause yields execution to OS
                
                ' Read progress updates from temp prog file
                If Dir(progFile) <> "" Then
                    On Error Resume Next
                    Dim fNum As Integer
                    fNum = FreeFile
                    Open progFile For Input As #fNum
                    Dim progLine As String
                    Line Input #fNum, progLine
                    Close #fNum
                    On Error GoTo ErrDB
                    
                    Dim slashPos As Long
                    slashPos = InStr(progLine, "/")
                    If slashPos > 0 Then
                        Dim currPage As Long
                        Dim totalPages As Long
                        currPage = Val(Left$(progLine, slashPos - 1))
                        totalPages = Val(Mid$(progLine, slashPos + 1))
                        
                        Call UpdateImportProgress(currPage, totalPages, startTime)
                    End If
                End If
                DoEvents
            Loop While exitCode = STILL_ACTIVE
            Call CloseHandle(hProcess)
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbDefault
    
    ' Hide progress components on exit
    picProgressOuter.Visible = False
    lblProgressText.Visible = False
    
    If Dir(tempFile) = "" Then
        Dim errLog As String
        Dim logLine As String
        Dim fNo As Integer
        
        errLog = ""
        If Dir(App.Path & "\parser_log.txt") <> "" Then
            fNo = FreeFile
            Open App.Path & "\parser_log.txt" For Input As #fNo
            Do While Not EOF(fNo) And Len(errLog) < 1500
                Line Input #fNo, logLine
                errLog = errLog & logLine & vbCrLf
            Loop
            Close #fNo
        End If
        
        If Trim$(errLog) <> "" Then
            MsgBox "Failed to parse PDF. Console Diagnostic Log:" & vbCrLf & vbCrLf & errLog, vbCritical, "Import Error"
        Else
            MsgBox "Failed to parse PDF. The parser executable failed to start or was blocked by antivirus.", vbCritical, "Import Error"
        End If
        Exit Sub
    End If
    

    ' 5. Verify extracted QR values against DB (Read-only validations)

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set ts = fso.OpenTextFile(tempFile, 1)

    

    duplicateFound = False

    Do While Not ts.AtEndOfStream

        qrValue = Trim$(ts.ReadLine)

        If Len(qrValue) > 0 Then

            ' Check database registries (Strictly read-only check)

            If Not IsValidForPrinting(qrValue, errMsg) Then

                MsgBox "Verification failed for code inside PDF: " & qrValue & vbCrLf & _
                       "Reason: " & errMsg & vbCrLf & _
                       "Import aborted to prevent double-printing.", vbCritical, "Duplicate Blocked"

                duplicateFound = True

                Exit Do

            Else

                colImportQRs.Add qrValue

            End If

        End If

    Loop

    ts.Close

    

    If duplicateFound Or colImportQRs.count = 0 Then

        On Error Resume Next

        Kill tempFile

        Exit Sub

    End If

    

    ' 6. Write PDF Entry and link QRs in In_QRmaster as Printed=0 (reserves them)

    On Error GoTo ErrDB

    Cn.BeginTrans

    

    Dim sql As String

    Dim pdfID As Long

    Dim rs As New ADODB.Recordset

    Dim fileName As String

    fileName = Mid$(pdfPath, InStrRev(pdfPath, "\") + 1)

    

    sql = "INSERT INTO In_CustomerPDF (ItemCode, FileName, FilePath, TotalQty, Printed, ReprintCount) VALUES (" & _
          "'" & Replace(itemCode, "'", "''") & "', '" & Replace(fileName, "'", "''") & "', " & _
          "'" & Replace(pdfPath, "'", "''") & "', " & colImportQRs.count & ", 0, 0)"

    Cn.Execute sql

    

    rs.Open "SELECT @@IDENTITY AS NextID", Cn, adOpenForwardOnly, adLockReadOnly

    pdfID = rs!NextID

    rs.Close

    

    Dim i As Long

    Dim qrMonth As String

    qrMonth = Format(Date, "mmyyyy")

    

    ' Insert reserved codes into In_QRmaster. Blocks future imports/generations.

    For i = 1 To colImportQRs.count

        sql = "INSERT INTO In_QRmaster (ItemCode, QRMonth, SerialNo, QRValue, GeneratedOn, Printed, IsCustomerQR, CustomerPDFID) VALUES (" & _
              "'" & Replace(itemCode, "'", "''") & "', '" & qrMonth & "', 0, '" & Replace(colImportQRs(i), "'", "''") & "', GETDATE(), 0, 1, " & pdfID & ")"

        Cn.Execute sql

    Next i

    

    Cn.CommitTrans

    MsgBox "Verification successful! Imported " & colImportQRs.count & " unique QR codes. Ready to print.", vbInformation

    Call LoadCustomerPDFGrid

    

    On Error Resume Next

    Kill tempFile

    Exit Sub

    

ErrDB:

    Cn.RollbackTrans

    MsgBox "Database error during import: " & Err.Description, vbCritical

    On Error Resume Next

    Kill tempFile

End Sub



' Load Imported PDFs Grid

Private Sub LoadCustomerPDFGrid()

    Dim rs As New ADODB.Recordset

    Dim r As Integer

    

    On Error Resume Next

    fgPDF.Clear

    fgPDF.Rows = 1

    fgPDF.TextMatrix(0, 0) = "ID"

    fgPDF.TextMatrix(0, 1) = "Model"

    fgPDF.TextMatrix(0, 2) = "File Name"

    fgPDF.TextMatrix(0, 3) = "Qty"

    fgPDF.TextMatrix(0, 4) = "Printed?"

    

    rs.Open "SELECT PDFID, ItemCode, FileName, TotalQty, Printed FROM In_CustomerPDF ORDER BY PDFID DESC", Cn, adOpenStatic, adLockReadOnly

    r = 1

    Do While Not rs.EOF

        fgPDF.Rows = fgPDF.Rows + 1

        fgPDF.TextMatrix(r, 0) = rs!PDFID

        fgPDF.TextMatrix(r, 1) = rs!ItemCode

        fgPDF.TextMatrix(r, 2) = rs!FileName

        fgPDF.TextMatrix(r, 3) = rs!TotalQty

        If rs!Printed = 1 Then

            fgPDF.TextMatrix(r, 4) = "Yes"

        Else

            fgPDF.TextMatrix(r, 4) = "No"

        End If

        rs.MoveNext

        r = r + 1

    Loop

    rs.Close

End Sub



' Print PDF with Reprint and Scan Checks

Private Sub cmdPrintPDF_Click()

    Dim selectedRow As Integer

    Dim pdfID As Long

    Dim fileName As String

    Dim filePath As String

    Dim printedStatus As String

    Dim supervisorID As String

    Dim success As Long
    Dim selectedPrinter As String
    Dim printFilePath As String
    Dim isTempFileCreated As Boolean
    Dim splitOutFile As String
    Dim splitCmd As String
    Dim splitResult As Long
    Dim originalPrinterName As String

    

    selectedRow = fgPDF.Row

    If selectedRow <= 0 Then

        MsgBox "Please select a Customer PDF row first.", vbExclamation

        Exit Sub

    End If

    

    pdfID = Val(fgPDF.TextMatrix(selectedRow, 0))

    fileName = fgPDF.TextMatrix(selectedRow, 2)

    printedStatus = fgPDF.TextMatrix(selectedRow, 4)

    

    ' 1. Retrieve full file path

    Dim rs As New ADODB.Recordset

    rs.Open "SELECT FilePath FROM In_CustomerPDF WHERE PDFID = " & pdfID, Cn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then

        filePath = rs!filePath

    End If

    rs.Close

    

    If Dir(filePath) = "" Then

        MsgBox "PDF file cannot be found at location: " & filePath, vbCritical

        Exit Sub

    End If

    

    ' 2. Intercept Reprint Guard

    If printedStatus = "Yes" Then

        ' Check permissions

        If Not GlobalCanReprint Then

            supervisorID = PromptForSupervisorAuth()

            If supervisorID = "" Then

                MsgBox "Reprint blocked. Supervisor authorization required.", vbCritical, "Access Denied"

                Exit Sub

            End If

        Else

            supervisorID = GlobalUserID

        End If

        

        ' Run extra scan double-check before reprinting

        If HasPDFBeenScanned(pdfID) Then

            Dim answer As VbMsgBoxResult

            answer = MsgBox("WARNING: One or more QR codes on this sheet have already been scanned on product boxes!" & vbCrLf & _
                            "Reprinting may result in double labeling. Proceed anyway?", vbYesNo + vbCritical, "Double Scan Warning")

            If answer = vbNo Then Exit Sub

        End If

        

        ' Audit Log Reprint action

        Call LogReprintAction(fileName, supervisorID)

    End If

    

    ' 3. Display standard Print Dialog and send PDF directly to the selected Printer


    selectedPrinter = DisplayPrintDlg()
    If selectedPrinter = "" Then Exit Sub ' Operator cancelled the dialog
    
    isTempFileCreated = False
    printFilePath = filePath
    
    If m_isRangeSelected Then
        ' User requested a specific page range! Extract pages using split_pdf.exe
        splitOutFile = App.Path & "\temp_print.pdf"
        
        If Dir(splitOutFile) <> "" Then
            On Error Resume Next
            Kill splitOutFile
            On Error GoTo ErrDB
        End If
        
        splitCmd = Chr$(34) & App.Path & "\split_pdf.exe" & Chr$(34) & " " & Chr$(34) & filePath & Chr$(34) & " " & Chr$(34) & splitOutFile & Chr$(34) & " " & m_fromPage & " " & m_toPage
        
        ' Run the split command synchronously using ShellAndWait
        splitResult = ShellAndWait(splitCmd, vbHide)
        
        If Dir(splitOutFile) <> "" Then
            printFilePath = splitOutFile
            isTempFileCreated = True
        Else
            MsgBox "Failed to extract PDF page range. Printing entire document instead.", vbExclamation
        End If
    End If
    originalPrinterName = Printer.DeviceName
    Call SetDefaultPrinter(selectedPrinter)
    success = ShellExecute(Me.hwnd, "print", printFilePath, "", "", SW_HIDE)
    
    ' Restore original default printer and clean up after delay
    Call Sleep(2000) ' Wait 2 seconds for printer spooler before restoring
    Call SetDefaultPrinter(originalPrinterName)
    
    If isTempFileCreated Then
        On Error Resume Next
        Kill printFilePath
        On Error GoTo ErrDB
    End If
    

    If success > 32 Then

        ' 4. Update Database statuses

        If printedStatus = "No" Then

            Cn.BeginTrans

            Cn.Execute "UPDATE In_CustomerPDF SET Printed = 1, PrintedOn = GETDATE() WHERE PDFID = " & pdfID

            ' Mark all linked individual QR codes in In_QRmaster as Printed=1

            Cn.Execute "UPDATE In_QRmaster SET Printed = 1 WHERE CustomerPDFID = " & pdfID

            Cn.CommitTrans

        Else

            Cn.Execute "UPDATE In_CustomerPDF SET ReprintCount = ReprintCount + 1 WHERE PDFID = " & pdfID

        End If

        

        MsgBox "Print request sent successfully to default printer!", vbInformation

        Call LoadCustomerPDFGrid

    Else
        MsgBox "Printing failed. Error Code: " & success, vbCritical
    End If
    Exit Sub
ErrDB:
    MsgBox "Error during printing: " & Err.Description, vbCritical, "Print Error"
End Sub



' Check if any QR code inside PDF ID has already been scanned in esjobscandata

Private Function HasPDFBeenScanned(pdfID As Long) As Boolean

    Dim rs As New ADODB.Recordset

    Dim sql As String

    sql = "SELECT COUNT(1) AS ScanCount FROM esjobscandata s " & _
          "INNER JOIN In_QRmaster q ON s.ItemDetail = q.QRValue " & _
          "WHERE q.CustomerPDFID = " & pdfID

          

    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    HasPDFBeenScanned = (rs!ScanCount > 0)

    rs.Close

End Function



' --- SUPERVISOR AUTHORIZATION & AUDIT HELPERS ---

Private Function PromptForSupervisorAuth() As String

    Dim username As String

    Dim password As String

    Dim rs As New ADODB.Recordset

    

    username = InputBox("Enter Supervisor Username:", "Supervisor Authorization Required")

    If Trim$(username) = "" Then

        PromptForSupervisorAuth = ""

        Exit Function

    End If

    

    password = InputBox("Enter Supervisor Password:", "Supervisor Authorization Required")

    If Trim$(password) = "" Then

        PromptForSupervisorAuth = ""

        Exit Function

    End If

    

    rs.Open "SELECT UserID FROM UserMaster WHERE UserID='" & Replace(username, "'", "''") & _
            "' AND PasswordHash='" & Replace(password, "'", "''") & "' AND (UserRole='Supervisor' OR UserRole='Admin')", Cn, adOpenForwardOnly, adLockReadOnly

            

    If Not rs.EOF Then

        PromptForSupervisorAuth = rs!UserID

    Else

        PromptForSupervisorAuth = ""

    End If

    rs.Close

End Function



Private Sub LogReprintAction(qrValue As String, supervisorID As String)

    On Error Resume Next

    Dim sql As String

    sql = "INSERT INTO QRReprintLog (QRValue, AuthorizedBy, ReprintedOn, Reason) " & _
          "VALUES ('" & Replace(qrValue, "'", "''") & "', '" & Replace(supervisorID, "'", "''") & "', GETDATE(), 'Reprint authorized by supervisor')"

    Cn.Execute sql

End Sub



' Reprint Section Grid Loading

Private Sub cmdLoadReprint_Click()

    Dim rs As New ADODB.Recordset

    Dim r As Integer

    Dim itemText As String

    Dim itemCode As String

    

    itemText = Trim$(cboItem.Text)

    If InStr(itemText, " - ") > 0 Then

        itemCode = Left$(itemText, InStr(itemText, " - ") - 1)

    Else

        itemCode = itemText

    End If

    

    On Error Resume Next

    fg.Clear

    fg.Rows = 1

    Call InitializeGrids

    

    rs.Open "SELECT ItemCode, QRValue, GeneratedOn, Printed FROM In_QRmaster WHERE ItemCode='" & Replace(itemCode, "'", "''") & "' AND IsCustomerQR = 0 ORDER BY QRID DESC", Cn, adOpenStatic, adLockReadOnly

    

    r = 1

    Do While Not rs.EOF

        fg.Rows = fg.Rows + 1

        fg.TextMatrix(r, 0) = r

        fg.TextMatrix(r, 1) = rs!ItemCode

        fg.TextMatrix(r, 2) = rs!qrvalue

        fg.TextMatrix(r, 3) = rs!generatedon

        

        ' Initialize row Select column as Wingdings empty checkbox (o)

        fg.Row = r

        fg.Col = 4

        fg.CellFontName = "Wingdings"

        fg.CellFontSize = 11

        fg.CellAlignment = 1 ' Left-aligned

        fg.TextMatrix(r, 4) = Chr$(111)

        

        rs.MoveNext

        r = r + 1

    Loop

    rs.Close

End Sub



' Toggle checkbox selection:

' - Clicking Row 0, Col 4 toggles Select All / Deselect All

' - Clicking any row toggles its checkmark and highlights it

Private Sub fg_Click()

    Dim r As Long, c As Long

    Dim clickedRow As Long, clickedCol As Long

    

    clickedRow = fg.MouseRow

    clickedCol = fg.MouseCol

    

    If fg.Rows <= 1 Then Exit Sub

    If clickedRow < 0 Or clickedCol < 0 Then Exit Sub

    

    If clickedRow = 0 And clickedCol = 4 Then

        ' Toggle Header Checkbox (Select All / Deselect All)

        fg.Redraw = False

        If fg.TextMatrix(0, 4) = Chr$(111) Then

            ' Check Header Checkbox

            fg.Row = 0

            fg.Col = 4

            fg.CellFontName = "Wingdings"

            fg.CellFontSize = 11

            fg.CellAlignment = 1

            fg.TextMatrix(0, 4) = Chr$(254) ' Checked

            

            ' Check all rows & highlight

            For r = 1 To fg.Rows - 1

                fg.TextMatrix(r, 4) = Chr$(254)

                For c = 0 To fg.Cols - 1

                    fg.Row = r

                    fg.Col = c

                    fg.CellBackColor = &HFFD9C2

                Next c

            Next r

        Else

            ' Uncheck Header Checkbox

            fg.Row = 0

            fg.Col = 4

            fg.CellFontName = "Wingdings"

            fg.CellFontSize = 11

            fg.CellAlignment = 1

            fg.TextMatrix(0, 4) = Chr$(111) ' Unchecked

            

            ' Uncheck all rows & reset colors

            For r = 1 To fg.Rows - 1

                fg.TextMatrix(r, 4) = Chr$(111)

                For c = 0 To fg.Cols - 1

                    fg.Row = r

                    fg.Col = c

                    fg.CellBackColor = vbWhite

                Next c

            Next r

        End If

        fg.Redraw = True

        

    ElseIf clickedRow > 0 Then

        ' Toggle individual row checkbox (toggles when clicking anywhere on the row)

        fg.Redraw = False

        

        If fg.TextMatrix(clickedRow, 4) = Chr$(111) Then

            ' Check individual row & highlight

            fg.TextMatrix(clickedRow, 4) = Chr$(254)

            For c = 0 To fg.Cols - 1

                fg.Row = clickedRow

                fg.Col = c

                fg.CellBackColor = &HFFD9C2

            Next c

        Else

            ' Uncheck individual row & reset colors

            fg.TextMatrix(clickedRow, 4) = Chr$(111)

            For c = 0 To fg.Cols - 1

                fg.Row = clickedRow

                fg.Col = c

                fg.CellBackColor = vbWhite

            Next c

        End If

        

        ' Update header checkbox dynamically: check it only if ALL rows are checked

        Dim allChecked As Boolean

        allChecked = True

        For r = 1 To fg.Rows - 1

            If fg.TextMatrix(r, 4) = Chr$(111) Then

                allChecked = False

                Exit For

            End If

        Next r

        

        fg.Row = 0

        fg.Col = 4

        fg.CellFontName = "Wingdings"

        fg.CellFontSize = 11

        fg.CellAlignment = 1

        If allChecked Then

            fg.TextMatrix(0, 4) = Chr$(254)

        Else

            fg.TextMatrix(0, 4) = Chr$(111)

        End If

        

        fg.Redraw = True

    End If

End Sub



' Reprint selected generated QR codes with authorization checks

Private Sub cmdReprint_Click()

    Dim r As Integer

    Dim total As Long

    Dim prn As Printer

    Dim found As Boolean

    Dim pages As Long, p As Long

    Dim i As Long, index As Long

    Dim labelSizeTw As Long, gapTw As Long, leftTw As Long, topTw As Long

    Dim qrpic As StdPicture

    Dim parts As Variant

    Dim qrvalue As String

    Dim p1 As String, p2 As String

    Dim txt As String

    Dim Qrprint() As String

    Dim needsAuth As Boolean

    Dim supervisorID As String

    

    total = 0

    needsAuth = False

    ReDim Qrprint(fg.Rows)

    

    ' Gather selected items from grid (using Wingdings checked square Chr$(254) in Col 4)

    For r = 1 To fg.Rows - 1

        If fg.TextMatrix(r, 4) = Chr$(254) Then

            Qrprint(total) = fg.TextMatrix(r, 2)

            needsAuth = True

            total = total + 1

        End If

    Next

    

    If total = 0 Then

        MsgBox "Please select QR codes from the grid first.", vbExclamation

        Exit Sub

    End If

    

    ' Enforce Role-based reprint credentials verification

    If needsAuth Then

        If Not GlobalCanReprint Then

            supervisorID = PromptForSupervisorAuth()

            If supervisorID = "" Then

                MsgBox "Reprint blocked. Supervisor authorization required.", vbCritical, "Access Denied"

                Exit Sub

            End If

        Else

            supervisorID = GlobalUserID

        End If

        

        ' Log supervisor approved reprint actions to audit table

        For i = 0 To total - 1

            Call LogReprintAction(Qrprint(i), supervisorID)

        Next i

    End If

    

    ' Display standard Print Dialog Box

    Dim selectedPrinter As String

    selectedPrinter = DisplayPrintDlg()

    If selectedPrinter = "" Then Exit Sub ' Operator cancelled the dialog

    

    For Each prn In Printers

        If prn.DeviceName = selectedPrinter Then

            Set Printer = prn

            found = True

            Exit For

        End If

    Next



    If Not found Then

        MsgBox "The selected printer is offline or not found.", vbCritical

        Exit Sub

    End If



    ' Sizing layouts

    labelSizeTw = MMToTwips(16)

    gapTw = MMToTwips(8)

    topTw = MMToTwips(2)

    Printer.ScaleMode = vbTwips

    piclayout.ScaleMode = vbTwips



    pages = (total + 3) \ 4



    ' Execute Reprint Job

    For p = 1 To pages

        Printer.Font.Name = "Calibri"

        Printer.Font.Size = 8

        Printer.FontBold = False



        For i = 0 To 3

            index = (p - 1) * 4 + i + 1

            If index <= total Then

                Set qrpic = GenerateQRPicture(Qrprint(index - 1))

                leftTw = MMToTwips(3) + i * (labelSizeTw + gapTw)



                If Not qrpic Is Nothing Then

                    Printer.PaintPicture qrpic, leftTw, topTw, labelSizeTw, labelSizeTw

                End If



                ' Parse and print text labels

                qrvalue = Qrprint(index - 1)

                parts = Split(qrvalue, "|")

                If UBound(parts) >= 6 Then

                    p1 = Trim(parts(2)) ' CodeStr

                    p2 = Trim(parts(6)) ' SerialNo

                    txt = p1 & "-" & p2

                Else

                    txt = qrvalue

                End If



                Printer.CurrentX = leftTw

                Printer.CurrentY = topTw + labelSizeTw + MMToTwips(1)



                Do While Printer.TextWidth(txt) > labelSizeTw And Printer.Font.Size > 5

                    Printer.Font.Size = Printer.Font.Size - 0.5

                Loop



                Printer.Print txt

            End If

        Next i



        If p < pages Then Printer.NewPage

    Next p

    Printer.EndDoc



    ' Update Reprint count and printed flag in database

    Cn.BeginTrans

    For i = 0 To total - 1

        Cn.Execute "UPDATE In_QRmaster SET ReprintCount = ISNULL(ReprintCount, 0) + 1, Printed = 1 WHERE QRValue = '" & Replace(Qrprint(i), "'", "''") & "'"

    Next i

    Cn.CommitTrans



    MsgBox "Reprint job sent successfully to: " & Printer.DeviceName, vbInformation

    

    ' Refresh last printed serial display

    Call UpdateLastPrintedSerial

End Sub



' Switch User / Logout

Private Sub cmdLogout_Click()

    Unload Me

    frmLogin.Show

End Sub



' Auto-filter ComboBox dynamically as the user types in Search box

Private Sub txtSearch_Change()

    Dim searchStr As String

    Dim i As Long

    Dim matchesCount As Long

    

    searchStr = Trim$(txtSearch.Text)

    

    ' Clear dropdown and filter from cached FullItemList

    cboItem.Clear

    

    matchesCount = 0

    For i = 0 To ItemCount - 1

        If Len(searchStr) = 0 Or InStr(1, UCase(FullItemList(i)), UCase(searchStr)) > 0 Then

            cboItem.AddItem FullItemList(i)

            matchesCount = matchesCount + 1

        End If

    Next i

    

    ' If matching items exist, display the dropdown list (do NOT set ListIndex as that closes the dropdown)

    If cboItem.ListCount > 0 Then

        Call SendMessage(cboItem.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

    End If

    

    ' Refresh last printed serial display

    Call UpdateLastPrintedSerial

End Sub



' Generate QR picture using zint.exe (creates temp BMP, returns StdPicture)

Private Function GenerateQRPicture(QRText As String) As StdPicture

    Dim tmpFile As String

    Dim cmd As String

    Dim success As Boolean

    

    tmpFile = App.Path & "\tmp_qr_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd * 1000)) & ".bmp"

    

    ' Build Command line for zint

    cmd = """" & ZINT_PATH & """ -o """ & tmpFile & """ -b 58 --scale=4 --data=""" & QRText & """"

    

    success = ShellAndWait(cmd, 5000)

    If Not success Then

        ' fallback: wait for file to appear

        If Not WaitForFile(tmpFile, 7000) Then

            On Error Resume Next

            Set GenerateQRPicture = Nothing

            Exit Function

        End If

    End If



    On Error GoTo ErrLoad

    Set GenerateQRPicture = LoadPicture(tmpFile)

    

    On Error Resume Next

    Kill tmpFile

    Exit Function



ErrLoad:

    Set GenerateQRPicture = Nothing

End Function



' Wait for file helper

Private Function WaitForFile(ByVal filePath As String, ByVal timeoutMs As Long) As Boolean

    Dim elapsed As Long

    elapsed = 0

    Do While Dir(filePath) = "" And elapsed < timeoutMs

        Sleep 100

        elapsed = elapsed + 100

        DoEvents

    Loop

    WaitForFile = (Dir(filePath) <> "")

End Function



' Display Windows Open File Dialog Box natively (requires no OCX dependencies)

Private Function ShowOpenFileDialog() As String

    Dim ofn As OPENFILENAME

    Dim filePath As String

    Dim ret As Long

    

    ofn.lStructSize = Len(ofn)

    ofn.hwndOwner = Me.hwnd

    

    ' PDF File Types Filter

    ofn.lpstrFilter = "PDF Files (*.pdf)" & Chr(0) & "*.pdf" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)

    

    ' Set up buffer for selected file path

    filePath = String$(512, 0)

    ofn.lpstrFile = filePath

    ofn.nMaxFile = Len(filePath)

    

    ofn.lpstrFileTitle = String$(512, 0)

    ofn.nMaxFileTitle = Len(ofn.lpstrFileTitle)

    

    ' Set initial directory lookup

    ofn.lpstrInitialDir = "D:\Files"

    ofn.lpstrTitle = "Select Customer PDF File"

    ofn.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST

    

    ret = GetOpenFileName(ofn)

    

    If ret <> 0 Then

        ' Trim trailing null characters from path

        ShowOpenFileDialog = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1)

    Else

        ShowOpenFileDialog = ""

    End If

End Function



' Execute selected Operational Report

Private Sub cmdRunReport_Click()

    Dim rs As New ADODB.Recordset

    Dim sql As String

    Dim r As Long

    

    Screen.MousePointer = vbHourglass

    fgReports.Clear

    fgReports.Rows = 1

    

    Select Case cboReportType.ListIndex

        Case 0 ' 1. Supervisor Reprint Audit Trail

            fgReports.Cols = 5

            fgReports.ColWidth(0) = 800

            fgReports.ColWidth(1) = 2200

            fgReports.ColWidth(2) = 5000

            fgReports.ColWidth(3) = 1500

            fgReports.ColWidth(4) = 1800

            

            fgReports.TextMatrix(0, 0) = "Row"

            fgReports.TextMatrix(0, 1) = "Reprinted On"

            fgReports.TextMatrix(0, 2) = "QR Value / PDF Filename"

            fgReports.TextMatrix(0, 3) = "Authorized By"

            fgReports.TextMatrix(0, 4) = "Reason"

            

            sql = "SELECT ReprintedOn, QRValue, AuthorizedBy, Reason FROM QRReprintLog ORDER BY LogID DESC"

            rs.Open sql, Cn, adOpenStatic, adLockReadOnly

            

            r = 1

            Do While Not rs.EOF

                fgReports.Rows = fgReports.Rows + 1

                fgReports.TextMatrix(r, 0) = r

                fgReports.TextMatrix(r, 1) = rs!ReprintedOn

                fgReports.TextMatrix(r, 2) = rs!QRValue

                fgReports.TextMatrix(r, 3) = rs!AuthorizedBy

                fgReports.TextMatrix(r, 4) = rs!Reason

                rs.MoveNext

                r = r + 1

            Loop

            rs.Close

            

        Case 1 ' 2. Registry vs Packing Scan Reconciliation

            fgReports.Cols = 6

            fgReports.ColWidth(0) = 600

            fgReports.ColWidth(1) = 1500

            fgReports.ColWidth(2) = 5000

            fgReports.ColWidth(3) = 2000

            fgReports.ColWidth(4) = 2200

            fgReports.ColWidth(5) = 2200

            

            fgReports.TextMatrix(0, 0) = "Row"

            fgReports.TextMatrix(0, 1) = "Model"

            fgReports.TextMatrix(0, 2) = "QR Code Value"

            fgReports.TextMatrix(0, 3) = "Registry Date"

            fgReports.TextMatrix(0, 4) = "Packaging Scan Date"

            fgReports.TextMatrix(0, 5) = "Current Status"

            

            ' Optimized using 30-day date window to prevent query timeout on unindexed production scan log

            sql = "SELECT q.ItemCode, q.QRValue, q.GeneratedOn, s.DATE AS ScanDate, " & _
                  "CASE WHEN s.ItemDetail IS NOT NULL THEN 'Scanned (Completed)' ELSE 'Unscanned (WIP)' END AS Status " & _
                  "FROM In_QRmaster q " & _
                  "LEFT JOIN esjobscandata s ON q.QRValue = s.ItemDetail AND s.DATE >= DATEADD(day, -30, GETDATE()) " & _
                  "WHERE q.GeneratedOn >= DATEADD(day, -30, GETDATE()) " & _
                  "ORDER BY q.QRID DESC"

            rs.Open sql, Cn, adOpenStatic, adLockReadOnly

            

            r = 1

            Do While Not rs.EOF

                fgReports.Rows = fgReports.Rows + 1

                fgReports.TextMatrix(r, 0) = r

                fgReports.TextMatrix(r, 1) = rs!ItemCode

                fgReports.TextMatrix(r, 2) = rs!QRValue

                fgReports.TextMatrix(r, 3) = rs!GeneratedOn

                fgReports.TextMatrix(r, 4) = IIf(IsNull(rs!ScanDate), "N/A - Not Scanned", rs!ScanDate)

                fgReports.TextMatrix(r, 5) = rs!Status

                rs.MoveNext

                r = r + 1

            Loop

            rs.Close

            

        Case 2 ' 3. Customer PDF Fulfillment Status

            fgReports.Cols = 8

            fgReports.ColWidth(0) = 600

            fgReports.ColWidth(1) = 800

            fgReports.ColWidth(2) = 1500

            fgReports.ColWidth(3) = 4200

            fgReports.ColWidth(4) = 800

            fgReports.ColWidth(5) = 1000

            fgReports.ColWidth(6) = 1800

            fgReports.ColWidth(7) = 1000

            

            fgReports.TextMatrix(0, 0) = "Row"

            fgReports.TextMatrix(0, 1) = "PDF ID"

            fgReports.TextMatrix(0, 2) = "Model"

            fgReports.TextMatrix(0, 3) = "Customer PDF Filename"

            fgReports.TextMatrix(0, 4) = "Qty"

            fgReports.TextMatrix(0, 5) = "Printed?"

            fgReports.TextMatrix(0, 6) = "Printed On"

            fgReports.TextMatrix(0, 7) = "Reprints"

            

            sql = "SELECT PDFID, ItemCode, FileName, TotalQty, Printed, PrintedOn, ReprintCount FROM In_CustomerPDF ORDER BY PDFID DESC"

            rs.Open sql, Cn, adOpenStatic, adLockReadOnly

            

            r = 1

            Do While Not rs.EOF

                fgReports.Rows = fgReports.Rows + 1

                fgReports.TextMatrix(r, 0) = r

                fgReports.TextMatrix(r, 1) = rs!PDFID

                fgReports.TextMatrix(r, 2) = rs!ItemCode

                fgReports.TextMatrix(r, 3) = rs!FileName

                fgReports.TextMatrix(r, 4) = rs!TotalQty

                fgReports.TextMatrix(r, 5) = IIf(rs!Printed = 1, "Yes", "No")

                fgReports.TextMatrix(r, 6) = IIf(IsNull(rs!PrintedOn), "N/A", rs!PrintedOn)

                fgReports.TextMatrix(r, 7) = rs!ReprintCount

                rs.MoveNext

                r = r + 1

            Loop

            rs.Close

            

        Case 3 ' 4. Daily Production Summary

            fgReports.Cols = 5

            fgReports.ColWidth(0) = 800

            fgReports.ColWidth(1) = 2200

            fgReports.ColWidth(2) = 2200

            fgReports.ColWidth(3) = 2800

            fgReports.ColWidth(4) = 2800

            

            fgReports.TextMatrix(0, 0) = "Row"

            fgReports.TextMatrix(0, 1) = "Date"

            fgReports.TextMatrix(0, 2) = "Model Code"

            fgReports.TextMatrix(0, 3) = "Total QR Codes Generated"

            fgReports.TextMatrix(0, 4) = "Total QR Codes Printed"

            

            ' Group by date only (ignoring time) and model code

            sql = "SELECT CONVERT(date, GeneratedOn) AS ProdDate, ItemCode, " & _
                  "COUNT(1) AS TotalGenerated, " & _
                  "SUM(CASE WHEN Printed = 1 THEN 1 ELSE 0 END) AS TotalPrinted " & _
                  "FROM In_QRmaster " & _
                  "GROUP BY CONVERT(date, GeneratedOn), ItemCode " & _
                  "ORDER BY ProdDate DESC, ItemCode"

            rs.Open sql, Cn, adOpenStatic, adLockReadOnly

            

            r = 1

            Do While Not rs.EOF

                fgReports.Rows = fgReports.Rows + 1

                fgReports.TextMatrix(r, 0) = r

                fgReports.TextMatrix(r, 1) = rs!ProdDate

                fgReports.TextMatrix(r, 2) = rs!ItemCode

                fgReports.TextMatrix(r, 3) = rs!TotalGenerated

                fgReports.TextMatrix(r, 4) = rs!TotalPrinted

                rs.MoveNext

                r = r + 1

            Loop

            rs.Close

    End Select

    

    Screen.MousePointer = vbDefault

    MsgBox "Report loaded successfully. " & (fgReports.Rows - 1) & " rows found.", vbInformation, "Success"

End Sub



' Export the current reports grid content to an Excel-compatible CSV file

Private Sub cmdExportCSV_Click()

    Dim r As Long, c As Long

    Dim line As String

    Dim fNo As Integer

    Dim filePath As String

    

    If fgReports.Rows <= 1 Then

        MsgBox "No report data available to export. Run a report first.", vbExclamation, "No Data"

        Exit Sub

    End If

    

    filePath = InputBox("Enter destination file path for CSV export:", "Export Report CSV", App.Path & "\Report_Output.csv")

    If Trim$(filePath) = "" Then Exit Sub

    

    On Error GoTo ErrExp

    fNo = FreeFile

    Open filePath For Output As #fNo

    

    ' Write headers and rows, wrapping values in quotes to handle commas

    For r = 0 To fgReports.Rows - 1

        line = ""

        For c = 0 To fgReports.Cols - 1

            line = line & """" & Replace(fgReports.TextMatrix(r, c), """", """""") & """"

            If c < fgReports.Cols - 1 Then line = line & ","

        Next c

        Print #fNo, line

    Next r

    Close #fNo

    

    MsgBox "Report successfully exported to CSV file!" & vbCrLf & filePath, vbInformation, "Export Successful"

    Exit Sub

    

ErrExp:

    MsgBox "CSV Export failed: " & Err.Description, vbCritical, "Export Error"

End Sub



' Display standard Windows Print Dialog Box natively via API (zero OCX dependencies)

Private Function DisplayPrintDlg() As String
    Dim pd As PRINTDLG_TYPE
    Dim pName As String
    
    pd.lStructSize = Len(pd)
    pd.hwndOwner = Me.hwnd
    
    ' Flags: Hide PrintToFile, Disable Selection, Enable Page Numbers
    pd.flags = &H100000 Or &H4
    
    pd.nMinPage = 1
    pd.nMaxPage = 9999
    pd.nFromPage = 1
    pd.nToPage = 1
    
    If PrintDlg(pd) <> 0 Then
        ' Retrieve selected printer name from memory handles
        pName = GetPrinterNameFromDevNames(pd.hDevNames)
        
        If (pd.flags And &H2) <> 0 Then
            m_isRangeSelected = True
            m_fromPage = pd.nFromPage
            m_toPage = pd.nToPage
        Else
            m_isRangeSelected = False
        End If
        
        ' Free memory allocated by the API
        If pd.hDevMode <> 0 Then Call GlobalFree(pd.hDevMode)
        If pd.hDevNames <> 0 Then Call GlobalFree(pd.hDevNames)
        
        DisplayPrintDlg = pName
    Else
        DisplayPrintDlg = ""
    End If
End Function



' Extract printer device name string from DEVNAMES handle

Private Function GetPrinterNameFromDevNames(ByVal hDevNames As Long) As String

    Dim pDevNames As Long

    Dim dn As DEVNAMES_TYPE

    Dim ch As Byte

    Dim i As Long

    Dim res As String

    

    pDevNames = GlobalLock(hDevNames)

    If pDevNames <> 0 Then

        Call CopyMemory(dn, ByVal pDevNames, Len(dn))

        

        res = ""

        i = 0

        Do

            Call CopyMemory(ch, ByVal (pDevNames + dn.wDeviceOffset + i), 1)

            If ch = 0 Then Exit Do

            res = res & Chr(ch)

            i = i + 1

        Loop While i < 260

        

        Call GlobalUnlock(hDevNames)

    End If

    GetPrinterNameFromDevNames = res

End Function







' Query and display the last generated QR details for the selected item (considering Brand 1612/Type 81 rules)

Private Sub UpdateLastPrintedSerial()

    Dim itemText As String

    Dim itemCode As String

    Dim sql As String

    Dim rsItem As New ADODB.Recordset

    Dim rsSerial As New ADODB.Recordset

    

    On Error GoTo ErrSerial

    

    itemText = Trim$(cboItem.Text)

    If itemText = "" Then

        lblLastSerial.Caption = "Last Generated QR: N/A"

        Exit Sub

    End If

    

    If InStr(itemText, " - ") > 0 Then

        itemCode = Left$(itemText, InStr(itemText, " - ") - 1)

    Else

        itemCode = itemText

    End If

    

    ' 1. Fetch brand and type

    sql = "SELECT ItemType, ItemBrand FROM ItemMaster WHERE CodeStr='" & Replace(itemCode, "'", "''") & "'"

    rsItem.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    If rsItem.EOF Then

        rsItem.Close

        lblLastSerial.Caption = "Last Generated QR: N/A"

        Exit Sub

    End If

    

    Dim itemType As Long

    Dim itemBrand As Long

    itemType = rsItem!ItemType

    itemBrand = Val(IIf(IsNull(rsItem!ItemBrand), 0, rsItem!ItemBrand))

    rsItem.Close

    

    ' 2. Query last generated QR details (regardless of printed status, ordered by most recent QRID DESC)

    If itemBrand = 1612 And itemType = 81 Then

        sql = "SELECT TOP 1 q.QRValue " & _
              "FROM In_QRmaster q " & _
              "INNER JOIN ItemMaster i ON q.ItemCode = i.CodeStr " & _
              "WHERE i.ItemBrand = 1612 AND i.ItemType = 81 " & _
              "ORDER BY q.QRID DESC"

    Else

        sql = "SELECT TOP 1 QRValue " & _
              "FROM In_QRmaster " & _
              "WHERE ItemCode = '" & Replace(itemCode, "'", "''") & "' " & _
              "ORDER BY QRID DESC"

    End If

    

    rsSerial.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    If Not rsSerial.EOF Then

        lblLastSerial.Caption = "Last Generated QR: " & rsSerial!QRValue

    Else

        lblLastSerial.Caption = "Last Generated QR: None"

    End If

    rsSerial.Close

    Exit Sub

    

ErrSerial:

    lblLastSerial.Caption = "Last Generated QR: Error"

End Sub



' Dynamically partition and distribute the grid width to all columns proportionally

Private Sub FitGridColumns()

    On Error Resume Next

    Dim w As Long

    

    ' 1. Distribute fg (Sequential Generation) columns: Row=5%, Model=15%, Details=52%, Created=18%, Select=10%

    w = fg.Width - 350

    If w > 1000 Then

        fg.ColWidth(0) = w * 0.05

        fg.ColWidth(1) = w * 0.15

        fg.ColWidth(2) = w * 0.52

        fg.ColWidth(3) = w * 0.18

        fg.ColWidth(4) = w * 0.1

    End If

    

    ' 2. Distribute fgPDF (Customer PDF) columns

    w = fgPDF.Width - 350

    If w > 1000 Then

        fgPDF.ColWidth(0) = w * 0.06

        fgPDF.ColWidth(1) = w * 0.16

        fgPDF.ColWidth(2) = w * 0.48

        fgPDF.ColWidth(3) = w * 0.1

        fgPDF.ColWidth(4) = w * 0.1

        fgPDF.ColWidth(5) = w * 0.05

        fgPDF.ColWidth(6) = w * 0.05

    End If

    

    ' 3. Distribute fgReports (Reports Dashboard) columns

    Call FitReportsGridColumns

End Sub



' Distribute Reports grid columns based on the selected report type

Private Sub FitReportsGridColumns()

    On Error Resume Next

    Dim w As Long

    w = fgReports.Width - 350

    If w <= 1000 Then Exit Sub

    

    Dim c As Long

    For c = 0 To fgReports.Cols - 1

        fgReports.ColAlignment(c) = 1 ' Default all data columns to Left-aligned

    Next c

    fgReports.ColAlignment(0) = 4 ' Keep row index centered

    

    Select Case cboReportType.ListIndex

        Case 0 ' 1. Supervisor Reprint Audit Trail: Row=5%, Date=20%, Details=45%, User=15%, Reason=15%

            fgReports.ColWidth(0) = w * 0.05

            fgReports.ColWidth(1) = w * 0.2

            fgReports.ColWidth(2) = w * 0.45

            fgReports.ColWidth(3) = w * 0.15

            fgReports.ColWidth(4) = w * 0.15

            

        Case 1 ' 2. Registry vs Packing Scan Reconciliation: Row=5%, Model=15%, Details=40%, Registry=15%, Scan=15%, Status=10%

            fgReports.ColWidth(0) = w * 0.05

            fgReports.ColWidth(1) = w * 0.15

            fgReports.ColWidth(2) = w * 0.4

            fgReports.ColWidth(3) = w * 0.15

            fgReports.ColWidth(4) = w * 0.15

            fgReports.ColWidth(5) = w * 0.1

            

        Case 2 ' 3. Customer PDF Fulfillment Status: Row=5%, ID=8%, Model=15%, Filename=38%, Qty=8%, Printed=8%, Date=10%, Reprints=8%

            fgReports.ColWidth(0) = w * 0.05

            fgReports.ColWidth(1) = w * 0.08

            fgReports.ColWidth(2) = w * 0.15

            fgReports.ColWidth(3) = w * 0.38

            fgReports.ColWidth(4) = w * 0.08

            fgReports.ColWidth(5) = w * 0.08

            fgReports.ColWidth(6) = w * 0.1

            fgReports.ColWidth(7) = w * 0.08

            

        Case 3 ' 4. Daily Production Summary: Row=10%, Date=25%, Model=25%, Generated=20%, Printed=20%

            fgReports.ColWidth(0) = w * 0.1

            fgReports.ColWidth(1) = w * 0.25

            fgReports.ColWidth(2) = w * 0.25

            fgReports.ColWidth(3) = w * 0.2

            fgReports.ColWidth(4) = w * 0.2

    End Select

End Sub



' Subroutine to calculate percentage completed andEstimated Time Left in real-time

Private Sub UpdateImportProgress(ByVal currPage As Long, ByVal totalPages As Long, ByVal startTime As Double)

    On Error Resume Next

    Dim pct As Double

    Dim elapsed As Double

    Dim estimatedTotal As Double

    Dim timeLeft As Double

    Dim timeStr As String

    

    If totalPages <= 0 Then Exit Sub

    

    pct = currPage / totalPages

    If pct > 1 Then pct = 1

    

    ' Update progress bar width

    picProgressInner.Width = picProgressOuter.ScaleWidth * pct

    

    ' Calculate elapsed time

    elapsed = Timer - startTime

    If elapsed < 0 Then elapsed = 0

    

    If currPage > 1 And elapsed > 0.5 Then

        estimatedTotal = elapsed / pct

        timeLeft = estimatedTotal - elapsed

        If timeLeft < 0 Then timeLeft = 0

        

        If timeLeft > 60 Then

            timeStr = Format(timeLeft \ 60, "0") & " min " & Format(Int(timeLeft Mod 60), "0") & " sec"

        Else

            timeStr = Format(Int(timeLeft), "0") & " sec"

        End If

    Else

        timeStr = "Calculating..."

    End If

    

    lblProgressText.Caption = "Parsing page " & currPage & " of " & totalPages & " (" & Format(pct * 100, "0") & "%) - Time Left: " & timeStr

    DoEvents

End Sub



' =========================================================================
' User Management Screen Implementation
' =========================================================================

' Load users into fgUsers Grid
Public Sub LoadUserGrid()
    Dim rs As New ADODB.Recordset
    Dim r As Integer
    Dim w As Long
    
    On Error GoTo ErrHandler
    
    fgUsers.Clear
    fgUsers.Rows = 1
    
    fgUsers.TextMatrix(0, 0) = "User ID"
    fgUsers.TextMatrix(0, 1) = "Password"
    fgUsers.TextMatrix(0, 2) = "Role"
    
    rs.Open "SELECT UserID, PasswordHash, UserRole FROM UserMaster ORDER BY UserID", Cn, adOpenStatic, adLockReadOnly
    
    r = 1
    Do While Not rs.EOF
        fgUsers.Rows = fgUsers.Rows + 1
        fgUsers.TextMatrix(r, 0) = rs!UserID
        fgUsers.TextMatrix(r, 1) = rs!PasswordHash
        fgUsers.TextMatrix(r, 2) = rs!UserRole
        rs.MoveNext
        r = r + 1
    Loop
    rs.Close
    
    ' Fit columns
    w = fgUsers.Width - 350
    If w > 0 Then
        fgUsers.ColWidth(0) = w * 0.35
        fgUsers.ColWidth(1) = w * 0.35
        fgUsers.ColWidth(2) = w * 0.3
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error loading users: " & Err.Description, vbCritical, "Database Error"
End Sub

' Selection Click Event for fgUsers
Private Sub fgUsers_Click()
    Dim selectedRow As Integer
    selectedRow = fgUsers.Row
    
    If selectedRow > 0 And selectedRow < fgUsers.Rows Then
        txtUserID.Text = fgUsers.TextMatrix(selectedRow, 0)
        txtPassword.Text = fgUsers.TextMatrix(selectedRow, 1)
        
        Dim role As String
        role = fgUsers.TextMatrix(selectedRow, 2)
        
        If role = "Operator" Then
            cboUserRole.ListIndex = 0
        ElseIf role = "Supervisor" Then
            cboUserRole.ListIndex = 1
        ElseIf role = "Admin" Then
            cboUserRole.ListIndex = 2
        Else
            cboUserRole.ListIndex = -1
        End If
        
        ' Lock UserID text box to prevent modifying user ID (primary key)
        txtUserID.Locked = True
        txtUserID.BackColor = &H00E0E0E0&
        
        ' Configure CRUD buttons for Edit/Delete mode
        cmdAddUser.Enabled = False
        cmdUpdateUser.Enabled = True
        cmdDeleteUser.Enabled = True
    End If
End Sub

' Clear UI user fields
Private Sub ClearUserFields()
    txtUserID.Text = ""
    txtPassword.Text = ""
    cboUserRole.ListIndex = -1
    txtUserID.Locked = False
    txtUserID.BackColor = &H00FFFFFF&
    
    ' Configure CRUD buttons for Add mode
    cmdAddUser.Enabled = True
    cmdUpdateUser.Enabled = False
    cmdDeleteUser.Enabled = False
End Sub

' Clear button click
Private Sub cmdClearFields_Click()
    Call ClearUserFields
End Sub

' Add User Button Click Event
Private Sub cmdAddUser_Click()
    Dim uID As String
    Dim uPass As String
    Dim uRole As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    uID = Trim$(txtUserID.Text)
    uPass = Trim$(txtPassword.Text)
    If cboUserRole.ListIndex >= 0 Then
        uRole = cboUserRole.List(cboUserRole.ListIndex)
    Else
        uRole = ""
    End If
    
    If uID = "" Or uPass = "" Or uRole = "" Then
        MsgBox "Please enter all details (User ID, Password, and Role).", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Check duplicate
    rs.Open "SELECT UserID FROM UserMaster WHERE UserID = '" & Replace(uID, "'", "''") & "'", Cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "User ID already exists. Please choose a different User ID.", vbCritical, "Duplicate User"
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    ' Insert user
    sql = "INSERT INTO UserMaster (UserID, PasswordHash, UserRole) VALUES (" & _
          "'" & Replace(uID, "'", "''") & "', " & _
          "'" & Replace(uPass, "'", "''") & "', " & _
          "'" & Replace(uRole, "'", "''") & "')"
    Cn.Execute sql
    
    MsgBox "User '" & uID & "' added successfully.", vbInformation, "Success"
    Call ClearUserFields
    Call LoadUserGrid
    Exit Sub
ErrHandler:
    MsgBox "Error adding user: " & Err.Description, vbCritical, "Database Error"
End Sub

' Update User Button Click Event
Private Sub cmdUpdateUser_Click()
    Dim uID As String
    Dim uPass As String
    Dim uRole As String
    Dim sql As String
    
    uID = Trim$(txtUserID.Text)
    uPass = Trim$(txtPassword.Text)
    If cboUserRole.ListIndex >= 0 Then
        uRole = cboUserRole.List(cboUserRole.ListIndex)
    Else
        uRole = ""
    End If
    
    If uID = "" Or uPass = "" Or uRole = "" Then
        MsgBox "Please select a user and fill in details.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Update user in DB
    sql = "UPDATE UserMaster SET " & _
          "PasswordHash = '" & Replace(uPass, "'", "''") & "', " & _
          "UserRole = '" & Replace(uRole, "'", "''") & "' " & _
          "WHERE UserID = '" & Replace(uID, "'", "''") & "'"
    Cn.Execute sql
    
    MsgBox "User '" & uID & "' updated successfully.", vbInformation, "Success"
    Call ClearUserFields
    Call LoadUserGrid
    Exit Sub
ErrHandler:
    MsgBox "Error updating user: " & Err.Description, vbCritical, "Database Error"
End Sub

' Delete User Button Click Event
Private Sub cmdDeleteUser_Click()
    Dim uID As String
    Dim sql As String
    Dim ans As VbMsgBoxResult
    
    uID = Trim$(txtUserID.Text)
    
    If uID = "" Then
        MsgBox "Please select a user to delete.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Safety check: Prevent Admin from deleting themselves
    If LCase(uID) = LCase(GlobalUserID) Then
        MsgBox "You cannot delete your own logged-in user account.", vbCritical, "Action Blocked"
        Exit Sub
    End If
    
    ans = MsgBox("Are you sure you want to delete user '" & uID & "'?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete")
    If ans = vbNo Then Exit Sub
    
    On Error GoTo ErrHandler
    
    ' Delete user from DB
    sql = "DELETE FROM UserMaster WHERE UserID = '" & Replace(uID, "'", "''") & "'"
    Cn.Execute sql
    
    MsgBox "User '" & uID & "' deleted successfully.", vbInformation, "Success"
    Call ClearUserFields
    Call LoadUserGrid
    Exit Sub
ErrHandler:
    MsgBox "Error deleting user: " & Err.Description, vbCritical, "Database Error"
End Sub

