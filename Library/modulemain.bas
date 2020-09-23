Attribute VB_Name = "mainmodule"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public view As Integer
Public fratepday As Integer
Public dayslimit As Integer
Public refcopy As Integer
Public maxhold As Integer
Public salnew As Integer
Public saltemp As Integer
Public salper As Integer
Public splashtime As Integer
Public welcometime As Integer
Public uname As String

Dim str As String
Dim modulers As ADODB.Recordset
Dim moduleconn As ADODB.Connection
Sub main()
  Set moduleconn = New ADODB.Connection
  moduleconn.CursorLocation = adUseClient
  moduleconn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

  Set modulers = New ADODB.Recordset
  modulers.Open "select Dayslimit,Fratepday,Maxhold,Pass,Refcopy,Salnew,Salper,Saltemp,Splashtime,Viewe,Welcometime from Custom", moduleconn, adOpenStatic, adLockOptimistic

            If modulers(3) = "administerpass" Then
                Load Frm_admin
                Frm_admin.Show
            Else
             Call globalload
            End If
End Sub
Public Sub globalload()
                view = modulers.Fields(9)
                fratepday = modulers.Fields(1)
                dayslimit = modulers.Fields(0)
                refcopy = modulers.Fields(4)
                maxhold = modulers.Fields(2)
                salnew = modulers.Fields(5)
                saltemp = modulers.Fields(7)
                salper = modulers.Fields(6)
                splashtime = modulers.Fields(8)
                welcometime = modulers.Fields(10)
                Adpass = modulers.Fields(3)
modulers.Close
Load frmSplash
frmSplash.Show
DoEvents

Sleep splashtime

Unload frmSplash
DoEvents
Load frmLogin
Load mdi_start
mdi_start.Show
frmLogin.Show

End Sub
Public Sub logoff()
  Unload frmLogin
  DoEvents
  Set moduleconn = New ADODB.Connection
  moduleconn.CursorLocation = adUseClient
  moduleconn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

  Set modulers = New ADODB.Recordset
  modulers.Open "select Dayslimit,Fratepday,Maxhold,Pass,Refcopy,Salnew,Salper,Saltemp,Splashtime,Viewe,Welcometime from Custom", moduleconn, adOpenStatic, adLockOptimistic
                
                view = modulers.Fields(9)
                fratepday = modulers.Fields(1)
                dayslimit = modulers.Fields(0)
                refcopy = modulers.Fields(4)
                maxhold = modulers.Fields(2)
                salnew = modulers.Fields(5)
                saltemp = modulers.Fields(7)
                salper = modulers.Fields(6)
                splashtime = modulers.Fields(8)
                welcometime = modulers.Fields(10)
modulers.Close
Load frmLogin
Load mdi_start
DoEvents
mdi_start.Show
frmLogin.Show
End Sub
