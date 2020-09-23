VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4050
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1488.898
   ScaleMode       =   0  'User
   ScaleWidth      =   3802.731
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_as 
      Height          =   315
      ItemData        =   "frmLogin.frx":24A2
      Left            =   1440
      List            =   "frmLogin.frx":24AC
      TabIndex        =   0
      ToolTipText     =   "Enter as"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txt_uname 
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
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "EmployeeID  for employee"
      Top             =   840
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Click to enter"
      Top             =   1920
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Click to abort"
      Top             =   1920
      Width           =   1140
   End
   Begin VB.TextBox txt_pass 
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
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Password"
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "&Enter as"
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
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   288
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
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
      TabIndex        =   5
      Top             =   872
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
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
      TabIndex        =   6
      Top             =   1354
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Loginrecord As ADODB.Recordset
Dim Loginconnection As ADODB.Connection

Private Sub cmb_as_Click()
If (cmb_as.Text = "Administrator") Then
txt_uname.Enabled = False
Else
txt_uname.Enabled = True
End If
txt_uname.Text = ""
End Sub

Private Sub cmdCancel_Click()
Unload mdi_start
End Sub
Private Sub cmdOK_Click()

If cmb_as.Text = "Administrator" Then
            If (txt_pass.Text = "") Then
            MsgBox "Please enter password.", vbInformation, "Password missing"
            txt_pass.SetFocus
            Exit Sub
            End If
  str = "Select Pass from Custom"
  Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
              If (txt_pass.Text = Loginrecord(0)) Then
               Loginrecord.Close
               uname = "Administrator"
               mdi_start.Enabled = True
               mdi_start.Show
               Me.Hide
               DoEvents
               Load Frm_welcome
               Frm_welcome.Show
               mdi_start.mnu_administer.Enabled = True
             Else
              Loginrecord.Close
              MsgBox "Invalid password.", vbInformation, "Acess Denied"
              txt_pass.Text = ""
              Exit Sub
             End If
ElseIf cmb_as.Text = "Employee" Then
            If (txt_uname.Text = "") Then
               MsgBox "Please enter User name.", vbInformation, "Username missing"
            ElseIf (txt_pass.Text = "") Then
               MsgBox "Please enter password.", vbInformation, "Password missing"
            Else
              str = "Select count(*) from Emptab where Empid = '" & Trim(txt_uname.Text) & "' and Psword = '" & Trim(txt_pass.Text) & "'"
              Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
                     If (Loginrecord(0) = 0) Then
                       MsgBox "Invalid password or Id.", vbInformation, "Acess Denied"
                       txt_uname.Text = ""
                       txt_pass.Text = ""
                       txt_uname.SetFocus
                       Loginrecord.Close
                       Exit Sub
                    Else
                Loginrecord.Close
                str = "Select Fname,Lname from Emptab where Empid = '" & Trim(txt_uname.Text) & "' and Psword = '" & Trim(txt_pass.Text) & "'"
                Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
                       uname = Loginrecord(0) & " " & Loginrecord(1)
                       mdi_start.Enabled = True
                       mdi_start.Show
                       Me.Hide
                       DoEvents
                       Load Frm_welcome
                       Frm_welcome.Show
                       Loginrecord.Close
                       mdi_start.mnu_administer.Enabled = False
                    End If
            End If
Else
MsgBox "Invalid enter Catagory.", vbCritical, "Invalid catagory"
End If
End Sub
Private Sub Form_Activate()
 cmb_as.SetFocus
 mdi_start.Enabled = False
End Sub

Private Sub Form_Load()
  On Error GoTo errlable
  Set Loginconnection = New ADODB.Connection
  Loginconnection.CursorLocation = adUseClient
  Loginconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
  
  str = "Select count(*) from Emptab"
  Set Loginrecord = New ADODB.Recordset
  Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
 If (Loginrecord(0) = 0) Then
 cmb_as.Text = "Administrator"
 cmb_as.Locked = True
 txt_uname.Enabled = False
 End If
Loginrecord.Close
Exit Sub

errlable:
MsgBox Err.Number & Err.Description
End Sub
