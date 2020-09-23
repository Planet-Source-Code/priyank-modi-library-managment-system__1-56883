VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm mdi_start 
   BackColor       =   &H8000000C&
   Caption         =   "Library managment system."
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "mdi_start.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdi_start.frx":0ECA
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9600
         TabIndex        =   7
         ToolTipText     =   "Calender"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19267585
         CurrentDate     =   38167
      End
      Begin VB.CommandButton cmd_search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         ToolTipText     =   "Click for search"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_return 
         Caption         =   "Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Click for return information"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_issue 
         Caption         =   "Issue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Click for issue informations"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   6
         ToolTipText     =   "End session"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_logoff 
         Caption         =   "Logoff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         ToolTipText     =   "Switch user"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_members 
         Caption         =   "Members"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Click for members details"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_books 
         Caption         =   "Books"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Click for Books detail"
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   10440
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16298
            Text            =   "Library managment system "
            TextSave        =   "Library managment system "
            Object.ToolTipText     =   "Graphics by Bhavesh modi, Contact : priyank_modi@yahoo.co.in"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "10/8/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "7:03 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4921
            Text            =   "Contact : priyank_modi@yahoo.co.in  "
            TextSave        =   "Contact : priyank_modi@yahoo.co.in  "
            Object.ToolTipText     =   "Created by : Priyank modi"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "mdi_start.frx":24863
   End
   Begin VB.Menu mnu_database 
      Caption         =   "&Database"
      Begin VB.Menu sm_books 
         Caption         =   "&Books"
         Shortcut        =   ^B
      End
      Begin VB.Menu sm_members 
         Caption         =   "&Members"
         Shortcut        =   ^M
      End
      Begin VB.Menu firstbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_logoff 
         Caption         =   "&Logoff"
         Shortcut        =   ^L
      End
      Begin VB.Menu firstbarsecond 
         Caption         =   "-"
      End
      Begin VB.Menu sm_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_tranjection 
      Caption         =   "T&ransaction"
      Begin VB.Menu sm_issue 
         Caption         =   "&Issue"
         Shortcut        =   ^I
      End
      Begin VB.Menu sm_return 
         Caption         =   "&Return"
         Shortcut        =   ^R
      End
      Begin VB.Menu secondbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_search 
         Caption         =   "&Search.."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnu_administer 
      Caption         =   "&Administrator"
      Begin VB.Menu sm_employees 
         Caption         =   "&Employees"
         Shortcut        =   ^E
      End
      Begin VB.Menu sm_global 
         Caption         =   "&Global"
         Shortcut        =   ^G
      End
      Begin VB.Menu thirdbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_backup 
         Caption         =   "Back up"
         Shortcut        =   ^U
      End
      Begin VB.Menu temp 
         Caption         =   "-"
      End
      Begin VB.Menu sm_settings 
         Caption         =   "Se&ttings"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "&Tools"
      Begin VB.Menu sm_notepad 
         Caption         =   "Notepad"
         Shortcut        =   ^N
      End
      Begin VB.Menu sm_calculator 
         Caption         =   "Calculator"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu sm_help 
         Caption         =   "Context"
         Shortcut        =   ^O
      End
      Begin VB.Menu sm_hsearch 
         Caption         =   "Search for help topic"
         Shortcut        =   ^H
      End
      Begin VB.Menu helpbar 
         Caption         =   "-"
      End
      Begin VB.Menu sm_about 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdi_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub cmd_books_Click()
Load Frm_books
Frm_books.Show
End Sub
Private Sub cmd_exit_Click()
Unload Me
End Sub
Private Sub cmd_issue_Click()
Load Frm_issue
Frm_issue.Show
End Sub
Private Sub cmd_logoff_Click()
If MsgBox("Are You Sure you want to logoff ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
Call logoff
DoEvents
End If
End Sub

Private Sub cmd_members_Click()
Load Frm_members
Frm_members.Show
End Sub

Private Sub cmd_Return_Click()
Load Frm_return
Frm_return.Show
End Sub

Private Sub cmd_search_Click()
Load Frm_search
Frm_search.Show
End Sub
Private Sub MDIForm_Load()
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height - 400
Me.Width = Screen.Width
Me.Show
Me.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
If MsgBox("Are You Sure you want to Quit ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
Unload frmLogin
Else
Cancel = True
End If
End Sub
Private Sub sbStatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
ShellExecute Me.hWnd, vbNullString, "http://www.yahoo.com", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub sm_about_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub sm_backup_Click()
Load Frm_backup
Frm_backup.Show
End Sub

Private Sub sm_books_Click()
Call cmd_books_Click
End Sub
Private Sub sm_calculator_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub
Private Sub sm_employees_Click()
Load Frm_Employees
Frm_Employees.Show
End Sub
Private Sub sm_exit_Click()
Call cmd_exit_Click
End Sub
Private Sub sm_global_Click()
Load Frm_global
Frm_global.Show
End Sub

Private Sub sm_help_Click()
 Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub sm_hsearch_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If


End Sub

Private Sub sm_issue_Click()
Call cmd_issue_Click
End Sub
Private Sub sm_logoff_Click()
Call cmd_logoff_Click
End Sub
Private Sub sm_members_Click()
Call cmd_members_Click
End Sub
Private Sub sm_notepad_Click()
On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\notepad.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub sm_return_Click()
Call cmd_Return_Click
End Sub
Private Sub sm_search_Click()
Call cmd_search_Click
End Sub
Private Sub sm_settings_Click()
Load Frm_settings
Frm_settings.Show
End Sub
