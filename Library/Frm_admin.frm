VERSION 5.00
Begin VB.Form Frm_admin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer password"
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "Frm_admin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1426.862
   ScaleMode       =   0  'User
   ScaleWidth      =   4013.994
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_pass1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Administrator password"
      Top             =   720
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Click to submit"
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Cancel to Abort"
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txt_pass2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Administrator password"
      Top             =   1200
      Width           =   2325
   End
   Begin VB.Label lbl_admin 
      Caption         =   "Enter password for Administrator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lbl_pass1 
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   762
      Width           =   1080
   End
   Begin VB.Label lbl_pass2 
      Caption         =   "Password &confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1244
      Width           =   1680
   End
End
Attribute VB_Name = "Frm_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Adconnnection As ADODB.Connection
Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub
Private Sub cmdOK_Click()
If txt_pass1.Text = "" Then
MsgBox "Please enter password.", vbInformation, "Password missing"
ElseIf txt_pass2.Text = "" Then
MsgBox "Please enter password confirm.", vbInformation, "Password missing"
ElseIf txt_pass1.Text <> txt_pass2.Text Then
MsgBox "May be typing mistake please verify the password.", vbInformation, "Password missing"
txt_pass2.Text = ""
txt_pass1.Text = ""
txt_pass1.SetFocus
Else
    If MsgBox("This password will be use as Administrator level security,Are you sure you want keep this password ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
     str = "UPDATE Custom SET Pass='" & Trim(txt_pass1.Text) & "'"
    Adconnnection.Execute str
    MsgBox "Adminster can configure library settings from menu Administrator/settings.", vbInformation, "Administrator settings"
    Call globalload
    DoEvents
    Me.Hide
    Unload Me
    Exit Sub
    Else
    txt_pass2.Text = ""
    txt_pass1.Text = ""
    txt_pass1.SetFocus
    Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo errlable
  Set Adconnnection = New ADODB.Connection
  Adconnnection.CursorLocation = adUseClient
  Adconnnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

Exit Sub
errlable:
MsgBox Err.Description & Err.Number
End Sub

