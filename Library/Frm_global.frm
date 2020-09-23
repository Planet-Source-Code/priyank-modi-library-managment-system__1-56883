VERSION 5.00
Begin VB.Form Frm_global 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library's global informations"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "Frm_global.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_amount 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txt_finem 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txt_typebook 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txt_investment 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txt_salary 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txt_tnemp 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txt_deposite 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txt_tnmem 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txt_issbooks 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txt_avabooks 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txt_tnbooks 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lbl_fine 
      Caption         =   "Total fine amount Rs."
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
      TabIndex        =   21
      Top             =   3900
      Width           =   2175
   End
   Begin VB.Label fine_m 
      Caption         =   "Total no. of  fine paid entry"
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
      TabIndex        =   20
      Top             =   3555
      Width           =   2655
   End
   Begin VB.Label lbl_type 
      Caption         =   "Total type of book"
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
      Top             =   645
      Width           =   2055
   End
   Begin VB.Label lbl_invest 
      Caption         =   "Total book's investments Rs."
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
      TabIndex        =   14
      Top             =   1740
      Width           =   2655
   End
   Begin VB.Label lbl_salary 
      Caption         =   "Employees monthly salary Rs."
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
      TabIndex        =   6
      Top             =   3195
      Width           =   2655
   End
   Begin VB.Label lbl_emp 
      Caption         =   "Total employees"
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
      TabIndex        =   5
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label lbl_deposite 
      Caption         =   "Total deposite Rs."
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
      TabIndex        =   4
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label lbl_tmem 
      Caption         =   "Total no. of members"
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
      TabIndex        =   3
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label lbl_isssue 
      Caption         =   "Books Issued"
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
      TabIndex        =   2
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Label lbl_ava 
      Caption         =   "Books available"
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
      TabIndex        =   1
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label lbl_books 
      Caption         =   "Total no. of books"
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
      TabIndex        =   0
      Top             =   280
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bookr As ADODB.Recordset
Dim empr As ADODB.Recordset
Dim memr As ADODB.Recordset
Dim finer As ADODB.Recordset
Dim Database As ADODB.Connection
Dim str As String
Private Sub Form_Load()
On Error GoTo errlable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
Set Database = New ADODB.Connection
Database.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
Call updatedata
Call showdata
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub showdata()
If memr.Fields(0) <> 0 Then
txt_tnmem.Text = memr.Fields(0)
txt_deposite.Text = memr.Fields(1)
Else
txt_tnmem.Text = 0
txt_deposite.Text = 0
End If

If empr.Fields(0) <> 0 Then
txt_tnemp.Text = empr.Fields(0)
txt_salary.Text = empr.Fields(1)
Else
txt_tnemp.Text = 0
txt_salary.Text = 0
End If

If bookr.Fields(4) <> 0 Then
txt_tnbooks.Text = bookr.Fields(0)
txt_avabooks.Text = bookr.Fields(1)
txt_issbooks.Text = bookr.Fields(2)
txt_investment.Text = bookr.Fields(3)
txt_typebook.Text = bookr.Fields(4)
Else
txt_tnbooks.Text = 0
txt_avabooks.Text = 0
txt_issbooks.Text = 0
txt_investment.Text = 0
txt_typebook.Text = 0
End If
If (finer.Fields(0) <> 0) Then
txt_finem.Text = finer.Fields(0)
txt_amount.Text = finer.Fields(1)
Else
txt_finem.Text = 0
txt_amount.Text = 0
End If
End Sub
Private Sub updatedata()
Set bookr = New ADODB.Recordset
str = "select sum(Totalno),sum(Avano),sum(Issno),sum(Price*Totalno),count(*) from Book"
bookr.Open str, Database, adOpenStatic, adLockOptimistic

Set memr = New ADODB.Recordset
str = "select count(*),sum(Deposite) from member"
memr.Open str, Database, adOpenStatic, adLockOptimistic

Set empr = New ADODB.Recordset
str = "select count(*),sum(Salary) from Emptab"
empr.Open str, Database, adOpenStatic, adLockOptimistic

Set finer = New ADODB.Recordset
str = "Select count(*),sum(Fine) from Fine"
finer.Open str, Database, adOpenStatic, adLockOptimistic

End Sub
