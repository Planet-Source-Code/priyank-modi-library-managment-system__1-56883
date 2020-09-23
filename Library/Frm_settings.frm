VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer settings"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "Frm_settings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7860
   Begin VB.CommandButton cmd_finedel 
      Caption         =   "Delete fine"
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
      TabIndex        =   18
      ToolTipText     =   "Format Fine info. database"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_default 
      Caption         =   "Default"
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
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Set Default settings"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_apply 
      Caption         =   "Apply"
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
      TabIndex        =   17
      ToolTipText     =   "Apply settings"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "OK"
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
      Left            =   5400
      TabIndex        =   16
      ToolTipText     =   "Ok"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Change 
      Caption         =   "Change"
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
      TabIndex        =   15
      ToolTipText     =   "Click to modify"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   12960
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   2280
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   538
      TabMaxWidth     =   2999
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Library"
      TabPicture(0)   =   "Frm_settings.frx":24A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_mem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Administrator"
      TabPicture(1)   =   "Frm_settings.frx":24BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_deletea"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_welcome"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_splash"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fra_form"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fra_pass"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbl_time"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbl_wl"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl_spl"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmd_deletea 
         Caption         =   "Delete All"
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
         Left            =   -69600
         MaskColor       =   &H8000000F&
         TabIndex        =   19
         ToolTipText     =   "Format the Database"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_welcome 
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_splash 
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Frame fra_form 
         Caption         =   "Open form"
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
         Left            =   -69720
         TabIndex        =   35
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opt_tl 
            Caption         =   "Top left"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton opt_ce 
            Caption         =   "Screen center"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton opt_def 
            Caption         =   "Default"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame fra_pass 
         Caption         =   "Password settings"
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
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   4935
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
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   360
            Width           =   2775
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
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lbl_p1 
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
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lbl_p2 
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
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Fra_mem 
         Caption         =   "Employee"
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
         Height          =   1935
         Left            =   3720
         TabIndex        =   27
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txt_per 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   6
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_temp 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txt_new 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   4
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lbl_per 
            Caption         =   "Permenent"
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
            TabIndex        =   31
            Top             =   1470
            Width           =   1215
         End
         Begin VB.Label lbl_temp 
            Caption         =   "Temporary"
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
            TabIndex        =   30
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label llbl_new 
            Caption         =   "Newly joined"
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
            TabIndex        =   29
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Employeees Salary settings"
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
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tranjection"
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
         Height          =   1935
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   3375
         Begin VB.TextBox txt_maxday 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   0
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_ref 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   3
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txt_fine 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   2
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txt_maxhold 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   1
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lbl_daylimit 
            Caption         =   "Max. days to hold book"
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
            TabIndex        =   26
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label lbl_refcopy 
            Caption         =   "Max.no of refrence"
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
            TabIndex        =   25
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label lbl_rate 
            Caption         =   "Fine charge per day"
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
            TabIndex        =   24
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Label lbl_maxbook 
            Caption         =   "Max. Books hold"
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
            TabIndex        =   23
            Top             =   750
            Width           =   1935
         End
      End
      Begin VB.Label Label2 
         Caption         =   "in ms"
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
         Left            =   -70560
         TabIndex        =   39
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_time 
         Caption         =   "in ms"
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
         Left            =   -70560
         TabIndex        =   38
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label lbl_wl 
         Caption         =   "Welcome screen stay time"
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
         Left            =   -74640
         TabIndex        =   37
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lbl_spl 
         Caption         =   "Splash screen stay time"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   1800
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frm_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Dim str As String
Dim ps As String
Dim rs As ADODB.Recordset
Dim db As ADODB.Connection

Private Sub cmd_apply_Click()
On Error GoTo errlable
If (cheak = True) Then
        If (opt_tl.Value = True) Then
        temp = 1
        ElseIf (opt_ce.Value = True) Then
        temp = 2
        Else
        temp = 3
        End If
                    str = " UPDATE Custom SET "
                    str = str & "Dayslimit = " & CDbl(txt_maxday.Text) & ", "
                    str = str & "Fratepday = " & CDbl(txt_fine.Text) & ", "
                    str = str & "Maxhold = " & CDbl(txt_maxhold.Text) & ", "
                    str = str & "Pass = '" & Trim(txt_pass1.Text) & "', "
                    str = str & "Refcopy = " & CDbl(txt_ref.Text) & ", "
                    str = str & "Salnew = " & CDbl(txt_new.Text) & ", "
                    str = str & "Salper = " & CDbl(txt_per.Text) & ", "
                    str = str & "Saltemp = " & CDbl(txt_temp.Text) & ", "
                    str = str & "Splashtime = " & CDbl(txt_splash.Text) & ", "
                    str = str & "Viewe = " & temp & ", "
                    str = str & "Welcometime = " & CDbl(txt_welcome.Text) & " WHERE Key=1"
        db.Execute str
        MsgBox "Changes are Applied.", vbInformation, "Save"
                        
                        cmd_Change.Enabled = True
                        cmd_deletea.Enabled = True
                        cmd_apply.Enabled = False
                        cmd_cancel.Caption = "OK"
                        Call locktext(True)
'Activate currently running variable with new value
                view = temp
                fratepday = CDbl(txt_fine.Text)
                dayslimit = CDbl(txt_maxday.Text)
                refcopy = CDbl(txt_ref.Text)
                maxhold = CDbl(txt_maxhold.Text)
                salnew = CDbl(txt_new.Text)
                saltemp = CDbl(txt_temp.Text)
                salper = CDbl(txt_per.Text)
                splashtime = CDbl(txt_splash.Text)
                welcometime = CDbl(txt_welcome.Text)

                If (temp = 1) Then
                opt_tl.Value = True
                ElseIf temp = 2 Then
                opt_ce.Value = True
                Else
                opt_def.Value = True
                End If
End If
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
              If (txt_fine.Text = "") Then
              MsgBox "Please enter fine amount.", vbInformation, "Field missing"
              ElseIf txt_maxday.Text = "" Then
              MsgBox "Please enter max. value of days for bookhold.", vbInformation, "Field missing"
              ElseIf txt_ref.Text = "" Then
              MsgBox "Please enter no. for refcopy.", vbInformation, "Field missing"
              ElseIf txt_maxhold.Text = "" Then
              MsgBox "Please enter max no. copy hold by Member.", vbInformation, "Field missing"
              ElseIf txt_new.Text = "" Then
              MsgBox "Please enter Salary for newly joined.", vbInformation, "Field missing"
              ElseIf txt_temp.Text = "" Then
              MsgBox "Please enter Salary for temporarily working.", vbInformation, "Field missing"
              ElseIf txt_per.Text = "" Then
              MsgBox "Please enter Salary for permenently working.", vbInformation, "Field missing"
              ElseIf txt_splash.Text = "" Then
              MsgBox "Please enter splashscreen stay time in ms.", vbInformation, "Field missing"
              ElseIf txt_welcome.Text = "" Then
              MsgBox "Please enter Welcome screen stay time in ms.", vbInformation, "Field missing"
              ElseIf txt_pass1.Text = "" Then
              MsgBox "Please enter Password.", vbInformation, "Field missing"
              ElseIf txt_pass2.Text = "" Then
              MsgBox "Please enter Passwordconfirm.", vbInformation, "Field missing"
              ElseIf Not IsNumeric(txt_fine.Text) Then
              MsgBox "Fine amount mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_maxday.Text) Then
              MsgBox "Max. day of bookhold mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_ref.Text) Then
              MsgBox "Max no.of refrence copy mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_maxhold.Text) Then
              MsgBox "Max no.of bookhold by member mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_new.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_temp.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_per.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_splash.Text) Then
              MsgBox "Splash screen stay time mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_welcome.Text) Then
              MsgBox "Welcome screen stay time mustbe Numeric.", vbInformation, "Improper value"
              ElseIf txt_pass2.Text <> txt_pass1.Text Then
              MsgBox "May be typing mistake,plese verify the password.", vbCritical, "Invalid password"
              Else
              flag = True
              End If
   cheak = flag
End Function
Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_Change_Click()
                Call locktext(False)
                cmd_Change.Enabled = False
                cmd_deletea.Enabled = False
                cmd_apply.Enabled = True
                cmd_cancel.Caption = "Cancel"
                End Sub
Private Sub locktext(val As Boolean)
                txt_fine.Locked = val
                txt_maxday.Locked = val
                txt_ref.Locked = val
                txt_maxhold.Locked = val
                txt_new.Locked = val
                txt_temp.Locked = val
                txt_per.Locked = val
                txt_splash.Locked = val
                txt_welcome.Locked = val
                txt_pass1.Locked = val
                txt_pass2.Locked = val
                opt_tl.Enabled = Not val
                opt_ce.Enabled = Not val
                opt_def.Enabled = Not val
                  
End Sub

Private Sub cmd_default_Click()
On Error GoTo errlable
                    str = " UPDATE Custom SET "
                    str = str & "Dayslimit = 15,"
                    str = str & "Fratepday = 1,"
                    str = str & "Maxhold = 2,"
                    str = str & "Pass = '" & Trim(ps) & "', "
                    str = str & "Refcopy = 2,"
                    str = str & "Salnew = 2000,"
                    str = str & "Salper = 4500,"
                    str = str & "Saltemp = 3000,"
                    str = str & "Splashtime = 2000,"
                    str = str & "Viewe = 3,"
                    str = str & "Welcometime =1000   WHERE Key=1"
        db.Execute str
        Call showdata
        MsgBox "Default Changes are Applied.", vbInformation, "Save"
                        
                        cmd_Change.Enabled = True
                        cmd_deletea.Enabled = True
                        cmd_apply.Enabled = False
                        cmd_cancel.Caption = "OK"
                        Call locktext(True)

'Activate currently running variable with new value
                view = 3
                fratepday = CDbl(txt_fine.Text)
                dayslimit = CDbl(txt_maxday.Text)
                refcopy = CDbl(txt_ref.Text)
                maxhold = CDbl(txt_maxhold.Text)
                salnew = CDbl(txt_new.Text)
                saltemp = CDbl(txt_temp.Text)
                salper = CDbl(txt_per.Text)
                splashtime = CDbl(txt_splash.Text)
                welcometime = CDbl(txt_welcome.Text)

                If (view = 1) Then
                opt_tl.Value = True
                ElseIf view = 2 Then
                opt_ce.Value = True
                Else
                opt_def.Value = True
                End If

Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_deletea_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete all the information about Library database except admin. settings,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
If MsgBox("You will never be able to retrive information back,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbCritical, "Warning") = vbYes Then
   str = "DELETE FROM Book"
db.Execute str
   str = "DELETE FROM Member"
db.Execute str
   str = "DELETE FROM Issue"
db.Execute str
   str = "DELETE FROM fine"
db.Execute str
MsgBox "All entry except Administrator settings and employee information are deleted sucessfully.", vbInformation, "Database formatted"
End If
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description

End Sub
Private Sub cmd_finedel_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete all the Fine information,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM fine"
   db.Execute str
MsgBox "Fine database entry deleted successfully.", vbInformation, "Delete"
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub

Private Sub Form_Load()
 On erro GoTo errlable
      If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If

 Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

  Set rs = New ADODB.Recordset
  rs.Open "select Dayslimit,Fratepday,Maxhold,Pass,Refcopy,Salnew,Salper,Saltemp,Splashtime,Viewe,Welcometime from Custom", db, adOpenStatic, adLockOptimistic
ps = rs.Fields(3)
cmd_cancel.Caption = "OK"
cmd_apply.Enabled = False
Call showdata
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub showdata()
           If rs.EOF = False And rs.BOF = False Then
                temp = rs.Fields(9)
                txt_fine.Text = rs.Fields(1)
                txt_maxday.Text = rs.Fields(0)
                txt_ref.Text = rs.Fields(4)
                txt_maxhold.Text = rs.Fields(2)
                txt_new.Text = rs.Fields(5)
                txt_temp.Text = rs.Fields(7)
                txt_per.Text = rs.Fields(6)
                txt_splash.Text = rs.Fields(8)
                txt_welcome.Text = rs.Fields(10)
                txt_pass1.Text = rs.Fields(3)
                txt_pass2.Text = rs.Fields(3)
           If (temp = 1) Then
           opt_tl.Value = True
           ElseIf temp = 2 Then
           opt_ce.Value = True
           Else
           opt_def.Value = True
           End If
End If
End Sub

