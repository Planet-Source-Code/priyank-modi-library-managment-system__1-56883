VERSION 5.00
Begin VB.Form Frm_books 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Detail"
   ClientHeight    =   5655
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7935
   Icon            =   "Frm_books.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_detail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   43
      ToolTipText     =   "Click to see all books"
      Top             =   5280
      Width           =   7455
   End
   Begin VB.Frame Fra_self 
      Caption         =   "Copy info."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1575
      Left            =   4680
      TabIndex        =   28
      Top             =   960
      Width           =   3015
      Begin VB.TextBox txt_totalno 
         DataField       =   "Totalno"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txt_issue 
         DataField       =   "Issno"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   38
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_avano 
         DataField       =   "Avano"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl_c3 
         Caption         =   "Available"
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
         TabIndex        =   42
         Top             =   400
         Width           =   855
      End
      Begin VB.Label lbl_c2 
         Caption         =   "Issued"
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
         TabIndex        =   41
         Top             =   765
         Width           =   615
      End
      Begin VB.Label lbl_c1 
         Caption         =   "Total copy"
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
         TabIndex        =   40
         Top             =   1130
         Width           =   975
      End
   End
   Begin VB.TextBox txt_title 
      DataField       =   "Title"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   27
      Top             =   240
      Width           =   6255
   End
   Begin VB.TextBox txt_publication 
      DataField       =   "Publication"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   26
      Top             =   600
      Width           =   6255
   End
   Begin VB.Frame Fra_Author 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1575
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txt_author1 
         DataField       =   "Author1"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_author2 
         DataField       =   "Author2"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_author3 
         DataField       =   "Author3"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lbl_a1 
         Caption         =   "First"
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
         TabIndex        =   25
         Top             =   400
         Width           =   495
      End
      Begin VB.Label lbl_a2 
         Caption         =   "Second"
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
         TabIndex        =   24
         Top             =   765
         Width           =   735
      End
      Begin VB.Label lbl_a3 
         Caption         =   "Third"
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
         TabIndex        =   23
         Top             =   1130
         Width           =   615
      End
   End
   Begin VB.TextBox txt_isbn 
      DataField       =   "ISBNNumber"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   18
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txt_price 
      DataField       =   "Price"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txt_subject 
      DataField       =   "Subject"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txt_pages 
      DataField       =   "Pages"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txt_edition 
      DataField       =   "Edition"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txt_Bookid 
      DataField       =   "Bookid"
      DataSource      =   "Adodc"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame frm_cmd 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   7455
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   240
         Picture         =   "Frm_books.frx":24A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move First"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   600
         Picture         =   "Frm_books.frx":27E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move Previous"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   3360
         Picture         =   "Frm_books.frx":2B26
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move Next"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   3720
         Picture         =   "Frm_books.frx":2E68
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move Last"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   345
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
         Left            =   2880
         TabIndex        =   7
         ToolTipText     =   "Delete record"
         Top             =   360
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Edit record"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H8000000B&
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
         Left            =   240
         MaskColor       =   &H00FF0000&
         Picture         =   "Frm_books.frx":31AA
         TabIndex        =   5
         ToolTipText     =   "Add new record"
         Top             =   360
         Width           =   1215
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
         Left            =   5880
         TabIndex        =   4
         ToolTipText     =   "Save record"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancel 
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
         TabIndex        =   3
         ToolTipText     =   "Cancel"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "&Reset"
         Enabled         =   0   'False
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
         TabIndex        =   2
         ToolTipText     =   "Reset fields"
         Top             =   360
         Width           =   1335
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
         Left            =   5880
         TabIndex        =   1
         ToolTipText     =   "Close"
         Top             =   840
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
         TabIndex        =   12
         Top             =   960
         Width           =   2400
      End
   End
   Begin VB.Label lbl_title 
      Caption         =   "Title "
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
      Left            =   600
      TabIndex        =   36
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lbl_pub 
      Caption         =   "Publication"
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
      TabIndex        =   35
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl_isbn 
      Caption         =   "ISBN no"
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
      Left            =   3600
      TabIndex        =   34
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lbl_price 
      Caption         =   "Price Rs."
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
      Left            =   720
      TabIndex        =   33
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lbl_subject 
      Caption         =   "Subject"
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
      Left            =   3600
      TabIndex        =   32
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbl_Pages 
      Caption         =   "Pages"
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
      Left            =   720
      TabIndex        =   31
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbl_edition 
      Caption         =   "Edition"
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
      Left            =   3600
      TabIndex        =   30
      Top             =   2640
      Width           =   735
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
      Left            =   720
      TabIndex        =   29
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "Frm_books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bookrecord As ADODB.Recordset
Dim Bookconnection As ADODB.Connection
 
Dim pos As Integer
Dim str As String
Dim slct As String
Dim saveflag As Boolean

'Function cheaking validity of textbox
Private Function cheak() As Boolean
'declaring variable
   Dim status As Boolean
   status = False
   
      If txt_title.Text = "" Then
          MsgBox ("Please enter the Title."), vbInformation, "Information required"
        ElseIf txt_publication.Text = "" Then
          MsgBox ("Please enter the Publications."), vbInformation, "Information required"
        ElseIf txt_author1.Text = "" Then
          MsgBox ("Please enter the  First Authors name."), vbInformation, "Information required"
        ElseIf txt_bookid.Text = "" Then
          MsgBox ("Please enter bookid distinct from other"), vbInformation, "Information required"
        ElseIf txt_pages.Text = "" Then
          MsgBox ("Please enter no of pages of book."), vbInformation, "Information required"
        ElseIf txt_price.Text = "" Then
          MsgBox ("Please enter the price."), vbInformation, "Information required"
        ElseIf txt_totalno.Text = "" Then
          MsgBox ("Please enter no of copies."), vbInformation, "Information required"
         ElseIf txt_issue.Text = "" Then
          MsgBox ("Please enter no of copies issued."), vbInformation, "Information required"
        ElseIf txt_avano.Text = "" Then
          MsgBox ("Please enter no of copies available."), vbInformation, "Information required"
        ElseIf txt_edition = "" Then
          MsgBox ("Please enter the detail about edition of book."), vbInformation, "Information required"
        ElseIf txt_subject.Text = "" Then
          MsgBox ("Please enter subject related to the book."), vbInformation, "Information required"
        ElseIf txt_isbn.Text = "" Then
          MsgBox ("Please enter ISBN no. for book."), vbInformation, "Information required"
        ElseIf IsNumeric(txt_author1.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf IsNumeric(txt_author2.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf IsNumeric(txt_author3.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_bookid.Text) Then
          MsgBox ("Bookid must be numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_pages.Text) Then
          MsgBox ("Enter page as in form of string of digits."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_price.Text) Then
          MsgBox ("Price must be digit form,enter valid price."), vbInformation, "Invalid information"
         ElseIf IsNumeric(txt_edition.Text) Then
          MsgBox ("Enter the valid string for edition."), vbInformation, "Invalid information"
         ElseIf IsNumeric(txt_subject.Text) Then
          MsgBox ("Subject name can not be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_totalno.Text) Then
         MsgBox ("Total no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_avano.Text) Then
         MsgBox ("Available no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_issue.Text) Then
         MsgBox ("Issue no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not (CDbl(txt_totalno.Text) = (CDbl(txt_avano.Text) + CDbl(txt_issue.Text))) Then
          MsgBox ("Possibly incorrect data in copy info. frame."), vbInformation, "Invalid information"
        Else
        status = True
        End If
   cheak = status
End Function
'subroutin for setting text box mode
Private Sub setlock(val As Boolean)
     txt_title.Locked = val
     txt_publication.Locked = val
     txt_author1.Locked = val
     txt_author2.Locked = val
     txt_author3.Locked = val
     txt_price.Locked = val
     txt_pages.Locked = val
     txt_subject.Locked = val
     txt_isbn.Locked = val
     txt_totalno.Locked = val
     txt_edition.Locked = val
     txt_bookid.Locked = val
     txt_issue.Locked = val
     txt_avano.Locked = val

End Sub
'make blank the text box
Private Sub clear()
            txt_title.Text = ""
            txt_publication.Text = ""
            txt_author1.Text = ""
            txt_author2.Text = ""
            txt_author3.Text = ""
            txt_price.Text = ""
            txt_subject.Text = ""
            txt_isbn.Text = ""
            txt_pages.Text = ""
            txt_totalno.Text = ""
            txt_avano.Text = ""
            txt_issue.Text = ""
            txt_edition.Text = ""
            txt_bookid.Text = ""

'set focus to fiRSt textbox
            txt_title.SetFocus
End Sub
Private Sub showdata()
  If Bookrecord.EOF = False And Bookrecord.BOF = False Then
          txt_author1.Text = Bookrecord.Fields(0)
          txt_author2.Text = Bookrecord.Fields(1)
          txt_author3.Text = Bookrecord.Fields(2)
          txt_avano.Text = Bookrecord.Fields(3)
          txt_bookid.Text = Bookrecord.Fields(4)
          txt_edition.Text = Bookrecord.Fields(5)
          txt_isbn.Text = Bookrecord.Fields(6)
          txt_issue.Text = Bookrecord.Fields(7)
          txt_pages.Text = Bookrecord.Fields(8)
          txt_price.Text = Bookrecord.Fields(9)
          txt_publication.Text = Bookrecord.Fields(10)
          txt_subject.Text = Bookrecord.Fields(11)
          txt_title.Text = Bookrecord.Fields(12)
          txt_totalno.Text = Bookrecord.Fields(13)
 End If
 End Sub
Private Sub setbutton(val As Boolean)
   cmdFirst.Enabled = val
    cmdPrevious.Enabled = val
    cmdNext.Enabled = val
    cmdLast.Enabled = val
    cmd_delete.Enabled = val
    cmd_edit.Enabled = val
    cmd_new.Enabled = val
    cmd_reset.Enabled = Not val
    cmd_save.Enabled = Not val
    cmdCancel.Enabled = Not val
End Sub

Private Sub cmd_close_Click()
Unload Me
'Load Frm_welcome
'Frm_welcome.Show
End Sub

Private Sub cmd_detail_Click()
Load Frm_bookd
Frm_bookd.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
On erro GoTo cancelerr
'disablink control
    setlock (True)
    lblStatus.Caption = " Cancel."
 
 If Bookrecord.BOF And Bookrecord.EOF Then
   GoTo newproc
 Else
   Bookrecord.MoveFirst
   Call showdata
 End If

newproc:
  txt_title.SetFocus
  Call setbutton(True)
Exit Sub
cancelerr:
MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Bookrecord.MoveFirst
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

   Bookrecord.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
 lblStatus.Caption = "               Move       >"
  
  If Not Bookrecord.EOF Then Bookrecord.MoveNext
  If Bookrecord.EOF And Bookrecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Bookrecord.MoveLast
    
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

  If Not Bookrecord.BOF Then Bookrecord.MovePrevious
  If Bookrecord.BOF And Bookrecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Bookrecord.MovePrevious
 
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
Private Sub Form_Load()
   On Error GoTo errlable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
   
   Set Bookconnection = New ADODB.Connection
  Bookconnection.CursorLocation = adUseClient
   Bookconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
   slct = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book Order by Bookid"
     Set Bookrecord = New ADODB.Recordset
     Bookrecord.Open slct, Bookconnection, adOpenStatic, adLockOptimistic
'show current record
 Call showdata
'disable buttons
  pos = Bookrecord.AbsolutePosition
  cmd_reset.Enabled = False
  cmd_save.Enabled = False
  cmdCancel.Enabled = False
 Exit Sub
errlable:
MsgBox Err.Number & Err.Description
 End Sub


Private Sub cmd_delete_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM Book WHERE "
   str = str & "Bookid = "
   str = str & CDbl(txt_bookid.Text)
   Bookconnection.Execute str
   Bookrecord.Requery
   MsgBox "Record deleted sucessfully.", vbinformayion, "Delete"

If Bookrecord.BOF And Bookrecord.EOF Then
    Call clear
    MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
    cmd_delete.Enabled = False
Else
   Bookrecord.MoveNext
      If Bookrecord.EOF Then
       Bookrecord.MoveLast
      End If
   Call showdata
End If

'message for status of mode
lblStatus.Caption = "Record deleted."
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmd_edit_Click()
On Error GoTo lable

'Make all entries in input mode
     Call setlock(False)
     saveflag = False
'message for status of mode
           lblStatus.Caption = " Edit record"
   Call setbutton(False)
  ' cmdcancel.Enabled = False
'set focus
            txt_title.SetFocus
Exit Sub
lable:
MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()
On Error GoTo lable

'Make all entries in input mode enable
     Call setlock(False)
 'clear the text field
     Call clear
    saveflag = True
    lblStatus.Caption = " Add new record."
       
Call setbutton(False)
Exit Sub
lable:
'Error handling statement
MsgBox Err.Description
End Sub

Private Sub cmd_reset_Click()
 'make all to input mode enable
           Call setlock(False)
           Call clear
End Sub

Private Sub cmd_save_Click()
On Error GoTo lable
'Make all entries in input mode enable
          Call setlock(False)
 'cheaking for validity condition
            If cheak = True Then
              If txt_author2.Text = "" Then
                txt_author2.Text = "None"
               End If
              If txt_author3.Text = "" Then
                txt_author3.Text = "None"
               End If
pos = Bookrecord.AbsolutePosition

'saving new record
If saveflag = True Then
str = "INSERT INTO Book"
str = str & "(Author1, Author2, Author3, Avano, Bookid, Edition, ISBNNumber, Issno, Pages, Price, Publication, Subject, Title, Totalno) "
str = str & "VALUES('" & Trim(txt_author1.Text) & "', "
str = str & "'" & Trim(txt_author2.Text) & "', "
str = str & "'" & Trim(txt_author3.Text) & "', "
str = str & CDbl(txt_avano.Text) & ", "
str = str & CDbl(txt_bookid.Text) & ", "
str = str & "'" & Trim(txt_edition.Text) & "', "
str = str & "'" & Trim(txt_isbn.Text) & "', "
str = str & CDbl(txt_issue.Text) & ", "
str = str & CDbl(txt_pages.Text) & ", "
str = str & CDbl(txt_price.Text) & ", "
str = str & "'" & Trim(txt_publication.Text) & "', "
str = str & "'" & Trim(txt_subject.Text) & "', "
str = str & "'" & Trim(txt_title.Text) & "', "
str = str & CDbl(txt_totalno.Text) & ")"
Bookconnection.Execute str
Else 'for editing the record
str = "UPDATE Book SET "
str = str & "Author1='" & Trim(txt_author1.Text) & "',"
str = str & "Author2='" & Trim(txt_author2.Text) & "',"
str = str & "Author3='" & Trim(txt_author3.Text) & "',"
str = str & "Avano=" & CDbl(txt_avano.Text) & ","
str = str & "Bookid=" & CDbl(txt_bookid.Text) & ","
str = str & "Edition='" & Trim(txt_edition.Text) & "',"
str = str & "ISBNNumber='" & Trim(txt_isbn.Text) & "',"
str = str & "Issno=" & CDbl(txt_issue.Text) & ","
str = str & "Pages=" & CDbl(txt_pages.Text) & ","
str = str & "Price=" & CDbl(txt_price.Text) & ","
str = str & "Publication='" & Trim(txt_publication.Text) & "',"
str = str & "Subject='" & Trim(txt_subject.Text) & "',"
str = str & "Title='" & Trim(txt_title.Text) & "',"
str = str & "Totalno=" & CDbl(txt_totalno.Text)
str = str & " WHERE Bookid=" & CDbl(txt_bookid.Text)
Bookconnection.Execute str
End If

'Make all entries input mode disable
Call setlock(True)

Bookrecord.Requery
Bookrecord.Move (pos - 1)
'show thw current data record
Call showdata
 'message for status of mode
           lblStatus.Caption = " New record Saved."
           MsgBox ("Record has been suceefully saved."), vbInformation, "Saving Record"
Call setbutton(True)
End If
Exit Sub
lable:
If Err.Number = -2147467259 Then
MsgBox ("BookID already exist,please enter anothe ID."), vbCritical, "BookID exist"
txt_bookid.SetFocus
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub


