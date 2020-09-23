VERSION 5.00
Begin VB.Form Frm_Employees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee's details"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000040&
   Icon            =   "Frm_Employee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_cmd 
      ForeColor       =   &H00000040&
      Height          =   1455
      Left            =   480
      TabIndex        =   38
      Top             =   4440
      Width           =   7935
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         Picture         =   "Frm_Employee.frx":24A2
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "19"
         ToolTipText     =   "Move First"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         Picture         =   "Frm_Employee.frx":27E4
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "20"
         ToolTipText     =   "Move Previous"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         Picture         =   "Frm_Employee.frx":2B26
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "21"
         ToolTipText     =   "Move Next"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3840
         Picture         =   "Frm_Employee.frx":2E68
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "22"
         ToolTipText     =   "Move Last"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmd_close 
         Caption         =   "Clo&se"
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Tag             =   "18"
         ToolTipText     =   "Close"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Tag             =   "17"
         ToolTipText     =   "Cancel"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Tag             =   "16"
         ToolTipText     =   "Save record"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Tag             =   "15"
         ToolTipText     =   "Reset fields"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Tag             =   "14"
         ToolTipText     =   "Delete record"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Tag             =   "13"
         ToolTipText     =   "Edit record"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Tag             =   "12"
         ToolTipText     =   "Add new record"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   960
         Width           =   2400
      End
   End
   Begin VB.Frame fra_personal 
      Caption         =   "Personal info"
      ForeColor       =   &H00000040&
      Height          =   2655
      Left            =   240
      TabIndex        =   30
      Top             =   1800
      Width           =   8415
      Begin VB.ComboBox cmb_sex 
         DataField       =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Frm_Employee.frx":31AA
         Left            =   2040
         List            =   "Frm_Employee.frx":31B4
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "7"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt_note 
         DataField       =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   8
         Tag             =   "8"
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox txt_phone 
         DataField       =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txt_mail 
         DataField       =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_add 
         DataField       =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1695
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   125
         MultiLine       =   -1  'True
         TabIndex        =   10
         Tag             =   "11"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_lname 
         DataField       =   "Lname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   4
         Tag             =   "4"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_fname 
         DataField       =   "Fname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   3
         Tag             =   "3"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbl_add 
         Caption         =   "Address"
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
         Left            =   4560
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbl_sex 
         Caption         =   "Sex"
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
         Left            =   480
         TabIndex        =   36
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label lbl_note 
         Caption         =   "Special note"
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
         Left            =   480
         TabIndex        =   35
         Top             =   2200
         Width           =   1215
      End
      Begin VB.Label lbl_phone 
         Caption         =   "Phone no."
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
         Left            =   480
         TabIndex        =   34
         Top             =   1490
         Width           =   975
      End
      Begin VB.Label lbl_mail 
         Caption         =   "Email address"
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
         Left            =   480
         TabIndex        =   33
         Top             =   1135
         Width           =   1335
      End
      Begin VB.Label lbl_lname 
         Caption         =   "Last name"
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
         Left            =   480
         TabIndex        =   32
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lbl_fname 
         Caption         =   "First name"
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
         Left            =   480
         TabIndex        =   31
         Top             =   425
         Width           =   1095
      End
   End
   Begin VB.Frame frm_post 
      Caption         =   "Office info."
      ForeColor       =   &H00000040&
      Height          =   1335
      Left            =   4920
      TabIndex        =   26
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cmb_post 
         DataField       =   "Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Frm_Employee.frx":31C6
         Left            =   1320
         List            =   "Frm_Employee.frx":31D3
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "9"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_sal 
         DataField       =   "Salary"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "10"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl_salary 
         Caption         =   "Salary"
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
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl_post 
         Caption         =   "Post-aid"
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
         Left            =   360
         TabIndex        =   28
         Top             =   415
         Width           =   855
      End
   End
   Begin VB.Frame fra_log 
      Caption         =   "Login info."
      ForeColor       =   &H00000040&
      Height          =   1575
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txt_pass2 
         DataField       =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_pass1 
         DataField       =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Tag             =   "1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_empid 
         DataField       =   "Empid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Password confirm"
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
         Left            =   360
         TabIndex        =   25
         Top             =   1100
         Width           =   1575
      End
      Begin VB.Label lbl_pass1 
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   24
         Top             =   757
         Width           =   975
      End
      Begin VB.Label lbl_ID 
         Caption         =   "Employee ID"
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
         Left            =   360
         TabIndex        =   23
         Top             =   415
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Emprecordset As ADODB.Recordset
Dim Empconnection As ADODB.Connection
Dim saveflag As Boolean
Dim str As String
Dim slct As String
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
                If txt_empid.Text = "" Then
                  MsgBox ("Please enter EmployeeID."), vbInformation, "Data missing"
                ElseIf txt_pass1.Text = "" Then
                  MsgBox ("Please enter Password."), vbInformation, "Data missing"
                ElseIf txt_pass2.Text = "" Then
                 MsgBox ("Please enter Password as Varifier so that wrong password can be detected."), vbInformation, "Data missing"
                ElseIf txt_fname.Text = "" Then
                 MsgBox ("Please enter Employee first name."), vbInformation, "Data missing"
                ElseIf txt_lname.Text = "" Then
                 MsgBox ("Please enter EmployeeID second name."), vbInformation, "Data missing"
                ElseIf cmb_sex.Text = "" Then
                  MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                ElseIf (cmb_sex.Text <> "Male" And cmb_sex.Text <> "Female") Then
                 MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                ElseIf cmb_post.Text = "" Then
                 MsgBox ("Please select the post-aid."), vbInformation, "Invalid arguments"
                ElseIf (cmb_post.Text <> "New" And cmb_post.Text <> "Temporary" And cmb_post.Text <> "Permanent") Then
                 MsgBox ("Please select the post-aid."), vbInformation, "Invalid arguments"
                ElseIf txt_add.Text = "" Then
                 MsgBox ("Please enter Employee contact address."), vbInformation, "Data missing"
                ElseIf txt_pass1.Text <> txt_pass2.Text Then
                 MsgBox ("May be typing mistake,Please re-enter the password."), vbInformation, "Invalid password"
                 txt_pass1.Text = ""
                 txt_pass2.Text = ""
                 txt_pass1.SetFocus
                Else
                 flag = True
                End If
cheak = flag
End Function
Private Sub showdata()
  If Emprecordset.EOF = False And Emprecordset.BOF = False Then
                    txt_add.Text = Emprecordset.Fields(0)
                    txt_mail.Text = Emprecordset.Fields(1)
                    txt_empid.Text = Emprecordset.Fields(2)
                    txt_fname.Text = Emprecordset.Fields(3)
                    txt_lname.Text = Emprecordset.Fields(4)
                    txt_phone.Text = Emprecordset.Fields(5)
                    cmb_post.Text = Emprecordset.Fields(6)
                    txt_pass1.Text = Emprecordset.Fields(7)
                    txt_pass2.Text = Emprecordset.Fields(7)
                    txt_sal.Text = Emprecordset.Fields(8)
                    cmb_sex.Text = Emprecordset.Fields(9)
                    txt_note.Text = Emprecordset.Fields(10)
                  
 End If
 End Sub
Private Sub clear()
                    txt_add.Text = ""
                    cmb_post.Text = ""
                    txt_mail.Text = ""
                    txt_empid.Text = ""
                    txt_fname.Text = ""
                    txt_lname.Text = ""
                    txt_note.Text = ""
                    txt_pass1.Text = ""
                    txt_pass2.Text = ""
                    txt_phone.Text = ""
                    txt_sal.Text = ""
                    cmb_sex.Text = ""
End Sub
Private Sub setlock(val As Boolean)
                    txt_add.Locked = val
                    cmb_post.Locked = val
                    txt_mail.Locked = val
                    txt_empid.Locked = val
                    txt_fname.Locked = val
                    txt_lname.Locked = val
                    txt_note.Locked = val
                    txt_pass1.Locked = val
                    txt_pass2.Locked = val
                    txt_phone.Locked = val
                    cmb_sex.Locked = val
End Sub
Private Sub button(val As Boolean)
                    cmd_new.Enabled = val
                    cmd_edit.Enabled = val
                    cmd_delete.Enabled = val
                    cmdFirst.Enabled = val
                    cmdLast.Enabled = val
                    cmdNext.Enabled = val
                    cmdPrevious.Enabled = val
                    cmd_save.Enabled = Not val
                    cmd_reset.Enabled = Not val
                    cmd_cancel.Enabled = Not val
                  
End Sub

Private Sub cmb_post_Click()
    If cmb_post.Text = "New" Then
        txt_sal.Text = salnew
     ElseIf cmb_post.Text = "Temporary" Then
        txt_sal.Text = saltemp
     Else
        txt_sal.Text = salper
     End If
   
End Sub

Private Sub cmd_cancel_Click()
On erro GoTo cancelerr
'disablink control
    setlock (True)
    lblStatus.Caption = " Cancel."
 
 If Emprecordset.BOF And Emprecordset.EOF Then
   GoTo newproc
 Else
   Emprecordset.MoveFirst
   Call showdata
 End If

newproc:
  txt_fname.SetFocus
'enable control
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    cmd_delete.Enabled = True
    cmd_edit.Enabled = True
    cmd_new.Enabled = True
'disable buttons
    cmd_reset.Enabled = False
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False

Exit Sub
cancelerr:
MsgBox Err.Description
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub Form_Load()
  On Error GoTo errlable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
  
  Set Empconnection = New ADODB.Connection
  Empconnection.CursorLocation = adUseClient
  Empconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
  
  slct = "select Address,Email,Empid,Fname,Lname,Phone,Pos,Psword,Salary,Sex,Spe from Emptab Order by Fname"
  Set Emprecordset = New ADODB.Recordset
  Emprecordset.Open slct, Empconnection, adOpenStatic, adLockOptimistic
 
 Call showdata
   cmd_reset.Enabled = False
   cmd_save.Enabled = False
   cmd_cancel.Enabled = False
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmd_delete_Click()
On Error GoTo delerr
 Beep
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
str = "DELETE FROM Emptab WHERE "
str = str & "Empid='"
str = str & Trim(txt_empid.Text) & "'"
'MsgBox str
Empconnection.Execute str
Emprecordset.Requery
MsgBox ("Record deleted Successfully."), vbInformation, "Delete"

        If Emprecordset.BOF And Emprecordset.EOF Then
            Call clear
            MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
            cmd_delete.Enabled = False
        Else
            Emprecordset.MoveNext
                If Emprecordset.EOF Then
                Emprecordset.MoveLast
                End If
            Call showdata
        End If
lblStatus.Caption = " Record deleted."
End If
Exit Sub
delerr:
MsgBox Err.Description
End Sub

Private Sub cmd_reset_Click()
Call clear
End Sub

Private Sub cmd_save_Click()
On erro GoTo saver
If cheak = True Then
  'Autocorrection procedure
      If cmb_post.Text = "New" Then
        txt_sal.Text = salnew
     ElseIf cmb_post.Text = "Temporary" Then
        txt_sal.Text = saltemp
     Else
        txt_sal.Text = salper
     End If
   
    If txt_mail.Text = "" Then
    txt_mail.Text = "None"
    End If
    
    If txt_phone.Text = "" Then
    txt_phone.Text = "None"
    End If
    
    If txt_note.Text = "" Then
    txt_note.Text = "None"
    End If
  
  pos = Emprecordset.AbsolutePosition
    
      If saveflag = True Then
            'for new record
            str = "INSERT INTO Emptab "
            str = str & "(Address,Email,Empid,Fname,Lname,Phone,Pos,Psword,Salary,Sex,Spe) "
            str = str & "VALUES" & "('" & Trim(txt_add.Text) & "', "
            str = str & "'" & Trim(txt_mail.Text) & "', "
            str = str & "'" & Trim(txt_empid.Text) & "', "
            str = str & "'" & Trim(txt_fname.Text) & "', "
            str = str & "'" & Trim(txt_lname.Text) & "', "
            str = str & "'" & Trim(txt_phone.Text) & "', "
            str = str & "'" & Trim(cmb_post.Text) & "', "
            str = str & "'" & Trim(txt_pass1.Text) & "', "
            str = str & CDbl(txt_sal.Text) & ","
            str = str & "'" & Trim(cmb_sex.Text) & "', "
            str = str & "'" & Trim(txt_note.Text) & "')"
            'MsgBox str
            Empconnection.Execute str, , adCmdText + adExecuteNoRecords
            
      Else
            'for updating current record
            str = "UPDATE Emptab SET "
            str = str & "Address = '" & Trim(txt_add.Text) & "',"
            str = str & " Pos = '" & Trim(cmb_post.Text) & "',"
            str = str & " Email = '" & Trim(txt_mail.Text) & "',"
            str = str & " Empid = '" & Trim(txt_empid.Text) & "',"
            str = str & " Fname = '" & Trim(txt_fname.Text) & "',"
            str = str & " Lname = '" & Trim(txt_lname.Text) & "',"
            str = str & " Spe = '" & Trim(txt_note.Text) & "',"
            str = str & " Psword = '" & Trim(txt_pass1.Text) & "',"
            str = str & " Phone = '" & Trim(txt_phone.Text) & "',"
            str = str & " Salary = " & CDbl(txt_sal.Text) & ","
            str = str & " Sex = '" & Trim(cmb_sex.Text) & "'"
            str = str & " WHERE Empid = '" & Trim(txt_empid.Text) & "'"
            'MsgBox str
            Empconnection.Execute str
            End If

            Emprecordset.Requery
            Emprecordset.Move (pos - 1)
            MsgBox ("Record saved successfully.")
            lblStatus.Caption = " Record saved"
            Call setlock(True)
            Call button(True)
            Call showdata
End If
Exit Sub
saver:
MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Emprecordset.MoveFirst
   lblStatus.Caption = "     <<      Move"
'show thw current data record
   Call showdata

Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
  lblStatus.Caption = "               Move       >>"

   Emprecordset.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
 lblStatus.Caption = "               Move       >"
  
  If Not Emprecordset.EOF Then Emprecordset.MoveNext
  If Emprecordset.EOF And Emprecordset.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Emprecordset.MoveLast
    
  End If
'show thw current data record
     Call showdata

Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
   lblStatus.Caption = "      <       Move"

  If Not Emprecordset.BOF Then Emprecordset.MovePrevious
  If Emprecordset.BOF And Emprecordset.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Emprecordset.MovePrevious
 
  End If
'show thw current data record
    Call showdata
Exit Sub

GoPrevError:
 If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub

Private Sub cmd_edit_Click()
On Error GoTo editerr
            'Call clear
            Call button(False)
            Call setlock(False)
            'cmd_cancel.Enabled = False
            saveflag = False
            lblStatus.Caption = " Edit record"
            txt_empid.Locked = True
            txt_fname.SetFocus
Exit Sub
editerr:
MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()
On Error GoTo newerr
            Call clear
            Call button(False)
            Call setlock(False)
            saveflag = True
            lblStatus.Caption = " Add new record."
            txt_empid.SetFocus
Exit Sub
newerr:
MsgBox Err.Description
End Sub

