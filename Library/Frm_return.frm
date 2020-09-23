VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_return 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "Frm_return.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4275
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
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmd_fine 
      Caption         =   "&Fine information"
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
      Left            =   480
      TabIndex        =   12
      ToolTipText     =   "Fine information"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmd_issue 
      Caption         =   "Switch to I&ssue"
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
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Switch to Issue"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmd_Return 
      Caption         =   "&Return"
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
      TabIndex        =   6
      ToolTipText     =   "Return book"
      Top             =   1800
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
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Add new"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txt_fine 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
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
   Begin VB.TextBox txt_bookid 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin MSMask.MaskEdBox msk_return 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Administrator default settings"
      Top             =   960
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
   Begin VB.Label lbl_fine 
      Caption         =   "Fine Rs."
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
      TabIndex        =   11
      Top             =   1350
      Width           =   735
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
      Left            =   480
      TabIndex        =   10
      Top             =   1005
      Width           =   1815
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
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   165
      Width           =   1095
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
      TabIndex        =   8
      Top             =   525
      Width           =   735
   End
End
Attribute VB_Name = "Frm_return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim amount As Integer
Dim str As String
Dim temp As ADODB.Recordset
Dim Returnconnection As ADODB.Connection
Private Sub setlock(val As Boolean)
msk_return.Enabled = Not val
txt_bookid.Locked = val
txt_memid.Locked = val
End Sub
Private Sub setbutton(val As Boolean)
cmd_add.Enabled = val
cmd_Return.Enabled = Not val
cmd_cancel.Enabled = Not val
End Sub
Private Sub cleartext()
msk_return.Text = "__/__/____"
txt_bookid.Text = ""
txt_memid.Text = ""
txt_fine.Text = ""
End Sub
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
If msk_return.Text = "__/__/____" Then
MsgBox "Please select the date.", vbInformation, "Field missing"
ElseIf txt_bookid.Text = "" Then
MsgBox "Please enter the Bookid.", vbInformation, "Field missing"
ElseIf txt_memid.Text = "" Then
MsgBox "Please enter the Memberid.", vbInformation, "Field missing"
Else
flag = True
End If
cheak = flag
End Function
Private Sub cmd_add_Click()
Call setlock(False)
Call setbutton(False)
Call cleartext
End Sub
Private Sub cmd_cancel_Click()
Call setlock(True)
Call cleartext
Call setbutton(True)
End Sub

Private Sub cmd_fine_Click()
Load Frm_Fine
Frm_Fine.Show
Unload Me
End Sub
Private Sub cmd_issue_Click()
Load Frm_issue
Frm_issue.Show
Unload Me
End Sub
Private Sub cmd_Return_Click()
On Error GoTo errlable
If (cheak = True) Then

'Search for return bookid and memid entry
str = "select count(*) from Issue where Memid = " & CDbl(txt_memid.Text) & " and Bookid = " & CDbl(txt_bookid.Text)
temp.Open str, Returnconnection, adOpenStatic, adLockOptimistic
            If (temp(0) = 0) Then
                    MsgBox "There is no such book issued for specified fields.", vbCritical, "Invalid arguments "
                    temp.Close
                    Call setlock(True)
                    Call setbutton(True)
                    Exit Sub
            End If
            temp.Close
'display info. & ask user for allow
If MsgBox("Return Info.:MemberId=" & CDbl(txt_memid.Text) & " And  BookId=" & CDbl(txt_bookid.Text), vbYesNo, "Confirm Data") = vbYes Then
  str = "select Areturndate,Bookid,Issuedate,Returndate,Memid from Issue where Memid = " & CDbl(txt_memid.Text) & " and Bookid = " & CDbl(txt_bookid.Text)
  temp.Open str, Returnconnection, adOpenStatic, adLockOptimistic
           amount = (Date - temp.Fields(3)) * fratepday
                
ignoreoverflow:
                If (amount < 0) Then
                  amount = 0  'convert negative amount to zero
                End If
          ' for amount case
                If (amount <= 0) Then
                    GoTo withoutfine    'submit book without fine
                ElseIf (amount > 0) Then
                'option for providing fine amount
                i = MsgBox("Members Total fine amount Rs : " & amount & " as per Rs : " & fratepday & " per Day charge.click yes if paying or click No if fine is collected from Members Deposite.", vbYesNoCancel + vbExclamation, "Confirm Data")
                    Select Case i
                    Case vbYes
                    Case vbNo
                    'transfer from deposite
                    str = "UPDATE Member SET Deposite = Deposite-" & CDbl(amount) & " WHERE Memid= " & Trim(txt_memid.Text)
                    Returnconnection.Execute str
                    MsgBox "The fine amount is transfer from members deposite.", vbInformation, "Fine"
                    Case vbCancel
                    'cancelling process of making entry
                    Call setlock(True)
                    Call setbutton(True)
                    MsgBox "Return process was cancelled.No more entry Updated.", vbInformation, "Fine"
                    Exit Sub
                    End Select
                        
                        'make entry in fine table
                        str = "INSERT INTO Fine (Areturndate,Bookid,Fine,Memid)"
                        str = str & "VALUES ('" & Format$(msk_return.Text, "mm/dd/yyyy") & "', "
                        str = str & CDbl(txt_bookid.Text) & ", "
                        str = str & CDbl(amount) & ", "
                        str = str & CDbl(txt_memid.Text) & ")"
                        Returnconnection.Execute str
                        
withoutfine:            'Update entry in Book table
                        str = "UPDATE Book SET "
                        str = str & "Avano = Avano+1,"
                        str = str & "Issno = Issno-1 WHERE Bookid = " & Trim(txt_bookid.Text)
                        Returnconnection.Execute str
                           
                        'Update entry in member table
                        str = "UPDATE Member SET "
                        str = str & "Bookinhand = Bookinhand-1 WHERE Memid = " & Trim(txt_memid.Text)
                        Returnconnection.Execute str
                           
                'delete entry from Issue table
                str = "DELETE * FROM Issue WHERE Memid = " & CDbl(txt_memid.Text) & " and Bookid = " & CDbl(txt_bookid.Text)
                Returnconnection.Execute str
                
                txt_fine.Text = amount
                
                MsgBox "fields entry Updated succesfully", vbInformation, "Book returned"
                
                End If
Else
Call setlock(True)
Call setbutton(True)
Exit Sub
End If
Call setlock(True)
Call setbutton(True)

End If 'validity check condition over
Exit Sub
errlable:
If (Err.Number = 6) Then
amount = 0
GoTo ignoreoverflow
ElseIf (Err.Number <> 0) Then
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

Set Returnconnection = New ADODB.Connection
Returnconnection.CursorLocation = adUseClient
Returnconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

Set temp = New ADODB.Recordset

txt_fine.Locked = True

Call setlock(True)
Call setbutton(True)
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub msk_return_GotFocus()
msk_return.Text = Format$(Now, "mm/dd/yyyy")
msk_return.Enabled = False
End Sub
Private Sub txt_fine_GotFocus()
MsgBox "Fine amount will be decided by itself.", vbInformation, "Self field propery"
End Sub
