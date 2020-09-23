VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_members 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member's Details"
   ClientHeight    =   5580
   ClientLeft      =   3765
   ClientTop       =   1740
   ClientWidth     =   8895
   Icon            =   "Frm_members.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flexgrid 
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   5640
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2778
      _Version        =   393216
      ForeColorSel    =   16777215
      GridColor       =   -2147483632
   End
   Begin VB.CommandButton cmd_books 
      BackColor       =   &H00000040&
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   25
      ToolTipText     =   "Click to see the books details contained by member"
      Top             =   5280
      Width           =   8415
   End
   Begin VB.Frame fra_cmd 
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
      Height          =   1215
      Left            =   600
      TabIndex        =   40
      Top             =   3960
      Width           =   7695
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "12"
         ToolTipText     =   "Add new record"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "&Edit"
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
         Left            =   1200
         TabIndex        =   14
         Tag             =   "13"
         ToolTipText     =   "Edit record"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
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
         Left            =   2280
         TabIndex        =   15
         Tag             =   "14"
         ToolTipText     =   "Delete record"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "&Reset"
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
         TabIndex        =   18
         Tag             =   "15"
         ToolTipText     =   "Reset fields"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Sa&ve"
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
         Left            =   3360
         TabIndex        =   16
         Tag             =   "16"
         ToolTipText     =   "Save record"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
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
         Left            =   4440
         TabIndex        =   17
         Tag             =   "17"
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_close 
         Caption         =   "Clo&se"
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
         Left            =   6600
         TabIndex        =   19
         Tag             =   "18"
         ToolTipText     =   "Close"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5400
         Picture         =   "Frm_members.frx":24A2
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "22"
         ToolTipText     =   "Move Last"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   5040
         Picture         =   "Frm_members.frx":27E4
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "21"
         ToolTipText     =   "Move Next"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   2280
         Picture         =   "Frm_members.frx":2B26
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "20"
         ToolTipText     =   "Move Previous"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   1920
         Picture         =   "Frm_members.frx":2E68
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "19"
         ToolTipText     =   "Move First"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   345
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
         Left            =   2640
         TabIndex        =   22
         Top             =   720
         Width           =   2400
      End
   End
   Begin VB.Frame Fra_library 
      Caption         =   "Library info."
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
      Height          =   1455
      Left            =   240
      TabIndex        =   33
      Top             =   2520
      Width           =   8415
      Begin MSMask.MaskEdBox msk_expr 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_join 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_memid 
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txt_note 
         DataField       =   "Note"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   9
         Tag             =   "8"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_bookhnd 
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_deposite 
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbl_mid 
         Caption         =   "Member ID"
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
         Left            =   5760
         TabIndex        =   42
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lbl_bookin 
         Caption         =   "Book in hand"
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
         Left            =   5160
         TabIndex        =   39
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lbl_depo 
         Caption         =   "Deposits"
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
         Left            =   5160
         TabIndex        =   38
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lbl_expr 
         Caption         =   "Date of expire(mm/dd/yyyy)"
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
         TabIndex        =   36
         Top             =   645
         Width           =   2535
      End
      Begin VB.Label lbl_join 
         Caption         =   "Date of join(mm/dd/yyyy)"
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
         TabIndex        =   35
         Top             =   300
         Width           =   2295
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
         Left            =   240
         TabIndex        =   34
         Top             =   1000
         Width           =   1215
      End
   End
   Begin VB.Frame fra_personal 
      Caption         =   "Personal info"
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
      Height          =   2295
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   8415
      Begin MSMask.MaskEdBox msk_bdate 
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_fname 
         DataField       =   "Fname"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   0
         Tag             =   "3"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txt_lname 
         DataField       =   "Lname"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "4"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_add 
         DataField       =   "Address"
         ForeColor       =   &H00400000&
         Height          =   1335
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   125
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "11"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_mail 
         DataField       =   "Email"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "5"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_phone 
         DataField       =   "Phone"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   3
         Tag             =   "6"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmb_sex 
         DataField       =   "Sex"
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Frm_members.frx":31AA
         Left            =   2040
         List            =   "Frm_members.frx":31B4
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "7"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lbl_birth 
         Caption         =   "Birthdate(mm/dd/yyyy)"
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
         Top             =   1840
         Width           =   1935
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
         TabIndex        =   32
         Top             =   425
         Width           =   1095
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
         TabIndex        =   31
         Top             =   780
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
         TabIndex        =   30
         Top             =   1135
         Width           =   1335
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
         TabIndex        =   29
         Top             =   1490
         Width           =   975
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
         TabIndex        =   28
         Top             =   1845
         Width           =   495
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
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frm_members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Memconnection As ADODB.Connection
Dim Memrecordset As ADODB.Recordset
Dim Flexgridset As ADODB.Recordset
Dim temp As ADODB.Recordset

Dim bookshow As Boolean
Dim saveflag As Boolean
Dim slct As String
Dim str As String
Dim pos As Integer
Private Sub clear()
                    txt_add.Text = ""
                    msk_bdate.Text = "__/__/____"
                    txt_bookhnd.Text = ""
                    txt_deposite.Text = ""
                    msk_expr.Text = "__/__/____"
                    msk_join.Text = "__/__/____"
                    txt_mail.Text = ""
                    txt_fname.Text = ""
                    txt_lname.Text = ""
                    txt_memid.Text = ""
                    txt_note.Text = ""
                    txt_phone.Text = ""
                    cmb_sex.Text = ""
End Sub
Private Sub locktext(val As Boolean)
                    txt_add.Locked = val
                    msk_bdate.Enabled = Not val
                    'txt_bookhnd.Locked = val
                    txt_deposite.Locked = val
                    msk_expr.Enabled = Not val
                    msk_join.Enabled = Not val
                    txt_mail.Locked = val
                    txt_fname.Locked = val
                    txt_lname.Locked = val
                    txt_memid.Locked = val
                    txt_note.Locked = val
                    txt_phone.Locked = val
                    cmb_sex.Locked = val
End Sub
Private Sub setbutton(val As Boolean)
               cmd_new.Enabled = val
               cmd_edit.Enabled = val
               cmd_delete.Enabled = val
               cmdFirst.Enabled = val
               cmdLast.Enabled = val
               cmdNext.Enabled = val
               cmdPrevious.Enabled = val
               cmd_cancel.Enabled = Not val
               cmd_save.Enabled = Not val
               cmd_reset.Enabled = Not val

End Sub
Private Function cheak() As Boolean
    Dim flag As Boolean
    flag = False
                 If txt_add.Text = "" Then
                MsgBox "Please enter member's address.", vbInformation, "Information required"
                 ElseIf msk_bdate.Text = "__/__/____" Then
                MsgBox "Please enter member's date of birth.", vbInformation, "Information required"
               '  ElseIf txt_bookhnd.Text = "" Then
               ' MsgBox "Please enter no of books contain by member.", vbInformation, "Information required"
                 ElseIf txt_deposite.Text = "" Then
                MsgBox "Please enter deposite amount.", vbInformation, "Information required"
                 ElseIf msk_expr.Text = "__/__/____" Then
                 MsgBox "Please enter date of account expire.", vbInformation, "Information required"
                ElseIf msk_join.Text = "__/__/____" Then
                MsgBox "Please enter date of join.", vbInformation, "Information required"
                 ElseIf txt_fname.Text = "" Then
                MsgBox "Please enter member's first name.", vbInformation, "Information required"
                 ElseIf txt_lname.Text = "" Then
                MsgBox "Please enter member's last name or family name.", vbInformation, "Information required"
                 ElseIf txt_memid.Text = "" Then
                MsgBox "Please enter member ID no.", vbInformation, "Information required"
                 ElseIf cmb_sex.Text = "" Then
                MsgBox "Please select sex.", vbInformation, "Information required"
                 ElseIf (cmb_sex.Text <> "Male" And cmb_sex.Text <> "Female") Then
                 MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                 ElseIf Not IsNumeric(txt_deposite.Text) Then
                 MsgBox ("Deposite must be Numeric value."), vbInformation, "Invalid arguments"
              '   ElseIf Not IsNumeric(txt_bookhnd.Text) Then
              '   MsgBox ("Book in hand must be Numeric."), vbInformation, "Invalid arguments"
                 ElseIf Not IsNumeric(txt_memid.Text) Then
                 MsgBox ("MemberID must be Numeric."), vbInformation, "Invalid arguments"
                 Else
                 flag = True
                End If
cheak = flag
End Function
Private Sub cmd_books_Click()
If (bookshow = True) Then
Me.Height = 5970
Else
Me.Height = 7755
End If
bookshow = Not bookshow
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmd_cancel_Click()
On erro GoTo cancelerr
'disablink control
    Call locktext(True)
    lblStatus.Caption = " Cancel."
 
 If Memrecordset.BOF And Memrecordset.EOF Then
   GoTo newproc
 Else
   Memrecordset.MoveFirst
   Call showdata
 End If

newproc:
  txt_fname.SetFocus
Call setbutton(True)
Exit Sub
cancelerr:
MsgBox Err.Description
End Sub

Private Sub cmd_delete_Click()
On erro GoTo lable
 Beep
str = "select Bookinhand from Member where Memid = " & CDbl(txt_memid.Text)
temp.Open str, Memconnection, adOpenStatic, adLockOptimistic
If temp(0) <> 0 Then
MsgBox "Member account cannot be deleeted because member has not returned books.", vbInformation, "Books not returned"
temp.Close
Exit Sub
End If
temp.Close
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM Member WHERE "
   str = str & "Memid = "
   str = str & CDbl(txt_memid.Text)
   Memconnection.Execute str
   Memrecordset.Requery
   MsgBox "Record deleted sucessfully.", vbinformayion, "Delete"

If Memrecordset.BOF And Memrecordset.EOF Then
    Call clear
    MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
    cmd_delete.Enabled = False
Else
   Memrecordset.MoveNext
      If Memrecordset.EOF Then
       Memrecordset.MoveLast
      End If
   Call showdata
End If

'message for status of mode
lblStatus.Caption = " Record deleted."
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmd_edit_Click()
Call locktext(False)
Call setbutton(False)
msk_bdate.Enabled = True
msk_expr.Enabled = True
msk_join.Enabled = True
txt_bookhnd.Locked = True
'cmd_cancel.Enabled = False
txt_fname.SetFocus
saveflag = False
lblStatus.Caption = " Edit record."
End Sub

Private Sub cmd_new_Click()
Call locktext(False)
Call clear
Call setbutton(False)
msk_bdate.Enabled = True
msk_expr.Enabled = True
msk_join.Enabled = True
txt_bookhnd.Text = 0
txt_fname.SetFocus
saveflag = True
lblStatus.Caption = " Add new record."
End Sub

Private Sub cmd_reset_Click()
Call clear
Call locktext(False)
End Sub

Private Sub cmd_save_Click()
'error cheaking and autocorrection handle
On Error GoTo errlable
If (cheak = True) Then
    If (txt_note.Text = "") Then
    txt_note.Text = "None"
    End If
    If (txt_phone.Text = "") Then
    txt_phone.Text = "None"
    End If
    If (txt_mail.Text = "") Then
    txt_mail.Text = "None"
    End If
    
If (saveflag = True) Then
           txt_bookhnd.Text = 0
            str = "INSERT INTO Member "
            str = str & "(Address, Birthdate, Bookinhand, Deposite, Doexpire, Dojoin, Email, Fname, Lname, Memid, Noted, Phone, Sex) "
            str = str & "VALUES('" & Trim(txt_add.Text) & "', "
            str = str & "'" & Trim(msk_bdate.Text) & "', "
            str = str & CDbl(txt_bookhnd.Text) & ", "
            str = str & CDbl(Trim(txt_deposite.Text)) & ", "
            str = str & "'" & Trim(msk_expr.Text) & "', "
            str = str & "'" & Trim(msk_join.Text) & "', "
            str = str & "'" & Trim(txt_mail.Text) & "', "
            str = str & "'" & Trim(txt_fname.Text) & "', "
            str = str & "'" & Trim(txt_lname.Text) & "', "
            str = str & CDbl(Trim(txt_memid.Text)) & ", "
            str = str & "'" & Trim(txt_note.Text) & "', "
            str = str & "'" & Trim(txt_phone.Text) & "', "
            str = str & "'" & Trim(cmb_sex.Text) & "' )"
            'MsgBox str
Memconnection.Execute str
Else
            str = "UPDATE Member SET "
            str = str & " Address = '" & Trim(txt_add.Text) & "',"
            str = str & " Birthdate  = '" & Trim(msk_bdate.Text) & "',"
            str = str & " Bookinhand = '" & Trim(txt_bookhnd.Text) & "',"
            str = str & " Deposite = " & CDbl(txt_deposite.Text) & ","
            str = str & " Doexpire = '" & Trim(msk_expr.Text) & "',"
            str = str & " Dojoin = '" & Trim(msk_join.Text) & "',"
            str = str & " Email = '" & Trim(txt_mail.Text) & "',"
            str = str & " Fname = '" & Trim(txt_fname.Text) & "',"
            str = str & " Lname = '" & Trim(txt_lname.Text) & "',"
            str = str & " Memid = " & CDbl(txt_memid.Text) & ","
            str = str & " Noted = '" & Trim(txt_note.Text) & "',"
            str = str & " Phone = '" & Trim(txt_phone.Text) & "',"
            str = str & " Sex = '" & Trim(cmb_sex.Text) & "'"
            str = str & " WHERE Memid= " & CDbl(txt_memid.Text)
            'MsgBox str
Memconnection.Execute str
End If

        Memrecordset.Requery
        Memrecordset.MoveFirst
        MsgBox ("Record saved successfully."), vbInformation, "Save"
        Call locktext(True)
        Call setbutton(True)
        Call showdata

End If
Exit Sub
errlable:
If (Err.Number = -2147467259) Then
MsgBox ("Member ID already exist,please enter anothe ID."), vbCritical, "MemberID exist"
txt_memid.SetFocus
ElseIf (Err.Number = -2147217913) Then
MsgBox ("May be date field pattern wrong."), vbCritical, "Date"
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
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
  Set Memconnection = New ADODB.Connection
  Memconnection.CursorLocation = adUseClient
  Memconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

  Set Memrecordset = New ADODB.Recordset
  Memrecordset.Open "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member Order by Memid", Memconnection, adOpenStatic, adLockOptimistic
   
 Set temp = New ADODB.Recordset
   Call showdata
  Set Flexgridset = New ADODB.Recordset
  Call flexupdate
   Call setbutton(True)
msk_bdate.Enabled = False
msk_expr.Enabled = False
msk_join.Enabled = False
bookshow = False
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub flexupdate()

 If Memrecordset.EOF = False And Memrecordset.BOF = False Then
 Flexgridset.Open "select count(*) from Book where Bookid in(select Bookid from Issue where Memid=" & Trim(txt_memid.Text) & ")", Memconnection, adOpenStatic, adLockOptimistic
 flexgrid.Rows = Flexgridset(0) + 2
 Flexgridset.Close

 Flexgridset.Open "select Author1,Author2,Author3,Bookid,Edition,ISBNNumber,Pages,Price,Publication,Subject,Title from Book where Bookid in(select Bookid from Issue where Memid=" & Trim(txt_memid.Text) & ")", Memconnection, adOpenStatic, adLockOptimistic
  flexgrid.Visible = True
  flexgrid.Cols = 11
 For pos = 0 To flexgrid.Rows - 1
      With flexgrid
        .TextMatrix(pos, 0) = ""
        .TextMatrix(pos, 1) = ""
        .TextMatrix(pos, 2) = ""
        .TextMatrix(pos, 3) = ""
        .TextMatrix(pos, 4) = ""
        .TextMatrix(pos, 5) = ""
        .TextMatrix(pos, 6) = ""
        .TextMatrix(pos, 7) = ""
        .TextMatrix(pos, 8) = ""
        .TextMatrix(pos, 9) = ""
        .TextMatrix(pos, 10) = ""
      End With
  Next pos
 pos = 0
 With flexgrid
    .FixedAlignment(1) = flexAlignCenterCenter
   
    .TextMatrix(0, 0) = "Bookid"
    .TextMatrix(0, 1) = "Title"
    .TextMatrix(0, 2) = "Author1"
    .TextMatrix(0, 3) = "Author2"
    .TextMatrix(0, 4) = "Author3"
    .TextMatrix(0, 5) = "Publication"
    .TextMatrix(0, 6) = "Edition"
    .TextMatrix(0, 7) = "Subject"
    .TextMatrix(0, 8) = "ISBN"
    .TextMatrix(0, 9) = "Price"
    .TextMatrix(0, 10) = "Pages"

    
    .ColWidth(0) = 700
    .ColWidth(1) = 2800
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 2300
    .ColWidth(6) = 1500
    .ColWidth(7) = 1700
    .ColWidth(8) = 1250
    .ColWidth(9) = 500
    .ColWidth(10) = 500
   
    While Not Flexgridset.EOF
        
        .TextMatrix(.Row, 0) = Flexgridset(3)
        .TextMatrix(.Row, 1) = Flexgridset(10)
        .TextMatrix(.Row, 2) = Flexgridset(0)
        .TextMatrix(.Row, 3) = Flexgridset(1)
        .TextMatrix(.Row, 4) = Flexgridset(2)
        .TextMatrix(.Row, 5) = Flexgridset(8)
        .TextMatrix(.Row, 6) = Flexgridset(4)
        .TextMatrix(.Row, 7) = Flexgridset(9)
        .TextMatrix(.Row, 8) = Flexgridset(5)
        .TextMatrix(.Row, 9) = Flexgridset(7)
        .TextMatrix(.Row, 10) = Flexgridset(6)
        .Row = .Row + 1
      
      Flexgridset.MoveNext
      Wend
    .Row = 1
 End With
Flexgridset.Close
End If
End Sub
Private Sub showdata()
  If Memrecordset.EOF = False And Memrecordset.BOF = False Then
                    txt_add.Text = Memrecordset.Fields(0)
                    msk_bdate.Text = Format$(Memrecordset.Fields(1), "MM/dd/yyyy")
                    txt_bookhnd.Text = Memrecordset.Fields(2)
                    txt_deposite.Text = Memrecordset.Fields(3)
                    msk_expr.Text = Format$(Memrecordset.Fields(4), "MM/dd/yyyy")
                    msk_join.Text = Format$(Memrecordset.Fields(5), "MM/dd/yyyy")
                    txt_mail.Text = Memrecordset.Fields(6)
                    txt_fname.Text = Memrecordset.Fields(7)
                    txt_lname.Text = Memrecordset.Fields(8)
                    txt_memid.Text = Memrecordset.Fields(9)
                    txt_note.Text = Memrecordset.Fields(10)
                    txt_phone.Text = Memrecordset.Fields(11)
                    cmb_sex.Text = Memrecordset.Fields(12)
 End If
 End Sub


Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Memrecordset.MoveFirst
   lblStatus.Caption = "      <<     Move"
'show thw current data record
   Call showdata
   Call flexupdate

Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
  lblStatus.Caption = "               Move       >>"

   Memrecordset.MoveLast
'show thw current data record
   Call showdata
   Call flexupdate

Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
Dim my As String
On Error GoTo GoNextError
 lblStatus.Caption = "               Move       >"
  
  If Not Memrecordset.EOF Then Memrecordset.MoveNext
  If Memrecordset.EOF And Memrecordset.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Memrecordset.MoveLast
    
  End If
'show thw current data record
     Call showdata
     Call flexupdate
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
   lblStatus.Caption = "      <       Move"

  If Not Memrecordset.BOF Then Memrecordset.MovePrevious
  If Memrecordset.BOF And Memrecordset.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Memrecordset.MovePrevious
 
  End If
'show thw current data record
    Call showdata
    Call flexupdate
Exit Sub

GoPrevError:
  If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub
