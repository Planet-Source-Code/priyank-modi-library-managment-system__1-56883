VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_issue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Frm_issue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
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
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Height          =   300
      Left            =   3600
      Picture         =   "Frm_issue.frx":24A2
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "22"
      ToolTipText     =   "Move Last"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Height          =   300
      Left            =   3240
      Picture         =   "Frm_issue.frx":27E4
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "21"
      ToolTipText     =   "Move Next"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   300
      Left            =   600
      Picture         =   "Frm_issue.frx":2B26
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "20"
      ToolTipText     =   "Move Previous"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   300
      Left            =   240
      Picture         =   "Frm_issue.frx":2E68
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "19"
      ToolTipText     =   "Move First"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmd_return 
      Caption         =   "Switch to &Return"
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
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Switch to Return form"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmd_issue 
      Caption         =   "I&ssue"
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
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Issue book"
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "&Add"
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
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Add new"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Fra_Date 
      Caption         =   "Date of"
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
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   3735
      Begin MSMask.MaskEdBox msk_return 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Administrator default settings"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   4194304
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_issue 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "Administrator default settings"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   4194304
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Doissue 
         Caption         =   "Issue (mm/dd/yyyy)"
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
         TabIndex        =   17
         Top             =   400
         Width           =   1815
      End
      Begin VB.Label lbl_Doreturn 
         Caption         =   "Return (mm/dd/yyyy)"
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
         TabIndex        =   16
         Top             =   750
         Width           =   1815
      End
   End
   Begin VB.TextBox txt_bookid 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txt_memid 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
      Left            =   960
      TabIndex        =   10
      Top             =   3120
      Width           =   2280
   End
   Begin VB.Label lbl_bookid 
      Caption         =   "Book ID"
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
      TabIndex        =   14
      Top             =   495
      Width           =   735
   End
   Begin VB.Label lbl_memberid 
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
      Left            =   480
      TabIndex        =   13
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim rmem As ADODB.Recordset
Dim rbook As ADODB.Recordset
Dim riss As ADODB.Recordset
Dim Issueconnection As ADODB.Connection
Dim Issuerecord As ADODB.Recordset
Private Sub cmd_add_Click()
Call cleartext
Call setbutton(False)
Call locktext(False)
End Sub
Private Sub locktext(val As Boolean)
txt_bookid.Locked = val
msk_issue.Enabled = Not val
msk_return.Enabled = Not val
txt_memid.Locked = val
End Sub
Private Sub setbutton(val As Boolean)
cmd_add.Enabled = val
cmd_Return.Enabled = val
cmdFirst.Enabled = val
cmdLast.Enabled = val
cmdNext.Enabled = val
cmdPrevious.Enabled = val
cmd_issue.Enabled = Not val
cmd_cancel.Enabled = Not val
End Sub
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
If msk_return.Text = "__/__/____" Then
MsgBox "Please select the date.", vbInformation, "Field missing"
ElseIf msk_issue.Text = "__/__/____" Then
ElseIf txt_bookid.Text = "" Then
MsgBox "Please enter the Bookid.", vbInformation, "Field missing"
ElseIf txt_memid.Text = "" Then
MsgBox "Please enter the Memberid.", vbInformation, "Field missing"
Else
flag = True
End If
cheak = flag
End Function
Private Sub cleartext()
txt_bookid.Text = ""
msk_issue.Text = "__/__/____"
msk_return.Text = "__/__/____"
txt_memid.Text = ""
End Sub
Private Sub cmd_cancel_Click()
Call locktext(True)
Call cleartext
Call setbutton(True)
End Sub
Private Sub cmd_issue_Click()
On Error GoTo errlable
If (cheak = True) Then

'If member id exists
str = "select count(*) from Member where Memid = " & Trim(txt_memid.Text)
rmem.Open str, Issueconnection, adOpenStatic, adLockOptimistic
If rmem(0) = 0 Then
    MsgBox ("Member with mentioned memberID does not exists."), vbCritical, "Invalid arguments"
    rmem.Close
    Exit Sub
Else
    'Is capable of holding book.
    rmem.Close
    str = "select Bookinhand from Member where Memid = " & Trim(txt_memid.Text)
    rmem.Open str, Issueconnection, adOpenStatic, adLockOptimistic
            If rmem(0) = maxhold Then
            MsgBox ("Members can not hold books greater then " & maxhold & "."), vbCritical, "Invalid arguments"
            rmem.Close
            GoTo recycle
            End If
End If
rmem.Close
'if book is present for specified bookid
str = "select count(*) from Book where Bookid = " & Trim(txt_bookid.Text)
rbook.Open str, Issueconnection, adOpenStatic, adLockOptimistic
If rbook(0) = 0 Then
    MsgBox ("Book with mentioned bookid does not exists."), vbCritical, "Invalid arguments"
    rbook.Close
    Exit Sub
Else
    'is there available copy
    rbook.Close
    str = "select Avano from Book where Bookid = " & Trim(txt_bookid.Text)
    rbook.Open str, Issueconnection, adOpenStatic, adLockOptimistic
            If rbook(0) <= refcopy Then
            MsgBox ("Book contains only refrence copies which cannot be issued."), vbCritical, "Invalid arguments"
            rbook.Close
            GoTo recycle
            End If
End If
rbook.Close
'member has same book or not
 str = "Select count(*) from Issue where Bookid = " & Trim(txt_bookid.Text) & " And Memid = " & Trim(txt_memid.Text)
 riss.Open str, Issueconnection, adOpenStatic, adLockOptimistic
 If (riss(0) <> 0) Then
     MsgBox ("Member has already issue mentioned book copy.member can not take same book again."), vbCritical, "Invalid arguments"
     riss.Close
 Exit Sub
 End If
 Beep
If MsgBox("Issue Info.:MemberId=" & CDbl(txt_memid.Text) & " And  BookId=" & CDbl(txt_bookid.Text), vbYesNo, "Confirm Data") = vbYes Then
            str = "INSERT INTO Issue"
            str = str & " (Areturndate,Bookid,Issuedate,Returndate,Memid) "
            str = str & "VALUES('" & Trim(msk_return.Text) & "', "
            str = str & CDbl(txt_bookid.Text) & ", "
            str = str & "'" & Trim(msk_issue.Text) & "', "
            str = str & "'" & Trim(msk_return.Text) & "', "
            str = str & CDbl(txt_memid.Text) & ")"
            Issueconnection.Execute str
            
            str = "UPDATE Book SET "
            str = str & "Avano = Avano-1,"
            str = str & "Issno = Issno+1 where Bookid = " & Trim(txt_bookid.Text)
            Issueconnection.Execute str
            
            str = "UPDATE Member SET "
            str = str & "Bookinhand = Bookinhand+1 where Memid = " & Trim(txt_memid.Text)
            Issueconnection.Execute str
            
            Issuerecord.Requery
            MsgBox "All entry Updated sucessfully.", vbInformation, "Record saved"
    Call locktext(True)
    Call setbutton(True)
Else
recycle:
    Call locktext(True)
    Call setbutton(True)
    Call cleartext
End If

End If
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_Return_Click()
Load Frm_return
Frm_return.Show
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo lable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
Set Issueconnection = New ADODB.Connection
Issueconnection.CursorLocation = adUseClient
 Issueconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

Set Issuerecord = New ADODB.Recordset
Issuerecord.Open "Select Areturndate,Bookid,Issuedate,Returndate,Memid from Issue Order by Memid", Issueconnection, adOpenStatic, adLockOptimistic

Set rmem = New ADODB.Recordset
Set rbook = New ADODB.Recordset
Set riss = New ADODB.Recordset

Call showdata
Call setbutton(True)
Call locktext(True)
Exit Sub

lable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub showdata()
If Issuerecord.EOF = False And Issuerecord.BOF = False Then
'msk_return.Text = Issuerecord.Fields(0)
txt_bookid.Text = Issuerecord.Fields(1)
msk_issue.Text = Format$(Issuerecord.Fields(2), "mm/dd/yyyy")
msk_return.Text = Format$(Issuerecord.Fields(3), "dd/mm/yyyy")
txt_memid.Text = Issuerecord.Fields(4)
End If
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Issuerecord.MoveFirst
   lblStatus.Caption = "      <<     Move"
'show thw current data record
   Call showdata

Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
  lblStatus.Caption = "               Move       >>"

   Issuerecord.MoveLast
'show thw current data record
   Call showdata

Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
Dim my As String
On Error GoTo GoNextError
 lblStatus.Caption = "               Move       >"
  
  If Not Issuerecord.EOF Then Issuerecord.MoveNext
  If Issuerecord.EOF And Issuerecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Issuerecord.MoveLast
    
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

  If Not Issuerecord.BOF Then Issuerecord.MovePrevious
  If Issuerecord.BOF And Issuerecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Issuerecord.MovePrevious
 
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
Private Sub msk_issue_GotFocus()
msk_issue.Text = Format$(Now, "mm/dd/yyyy")
msk_issue.Enabled = False
End Sub
Private Sub msk_return_GotFocus()
msk_return.Text = Format$(Now + dayslimit, "mm/dd/yyyy")
msk_return.Enabled = False
End Sub
