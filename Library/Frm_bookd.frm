VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_bookd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " All book's details"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   7440
   Icon            =   "Frm_bookd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdLast 
      Height          =   300
      Left            =   3720
      Picture         =   "Frm_bookd.frx":24A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Move Last"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Height          =   300
      Left            =   3360
      Picture         =   "Frm_bookd.frx":27E4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move Next"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   300
      Left            =   600
      Picture         =   "Frm_bookd.frx":2B26
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move Previous"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   300
      Left            =   240
      Picture         =   "Frm_bookd.frx":2E68
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move First"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "&Back to Main"
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
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Back to booksform"
      Top             =   3000
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid Datagrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Detail view of books"
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      TabIndex        =   6
      Top             =   3000
      Width           =   2400
   End
End
Attribute VB_Name = "Frm_bookd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Dtbookrecord As ADODB.Recordset
Dim Dtbookconn As ADODB.Connection
Private Sub cmd_back_Click()
Load Frm_books
Frm_books.Show
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
  Set Dtbookconn = New ADODB.Connection
  Dtbookconn.CursorLocation = adUseClient
  Dtbookconn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
  str = "select Bookid,Title,Author1,Author2,Author3,ISBNNumber,Pages,Price,Publication,Subject,Avano as Availableno,Edition,Issno as Issueno,Totalno from Book Order by Bookid"
  Set Dtbookrecord = New ADODB.Recordset
  Dtbookrecord.Open str, Dtbookconn, adOpenStatic, adLockOptimistic
            Datagrid.Visible = True
            Set Datagrid.DataSource = Dtbookrecord
            Datagrid.ReBind
  Exit Sub
errlable:
MsgBox Err.Number & Err.Description
 End Sub

Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Dtbookrecord.MoveFirst
   lblStatus.Caption = "     <<      Move"
'show thw current data record
Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
  lblStatus.Caption = "               Move       >>"

   Dtbookrecord.MoveLast
'show thw current data record
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
 lblStatus.Caption = "               Move       >"
  
  If Not Dtbookrecord.EOF Then Dtbookrecord.MoveNext
  If Dtbookrecord.EOF And Dtbookrecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Dtbookrecord.MoveLast
    
  End If
'show thw current data record

Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
   lblStatus.Caption = "      <       Move"

  If Not Dtbookrecord.BOF Then Dtbookrecord.MovePrevious
  If Dtbookrecord.BOF And Dtbookrecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Dtbookrecord.MovePrevious
 
  End If
'show thw current data record
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub

