VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmReport1 
   Caption         =   "REPORT"
   ClientHeight    =   6525
   ClientLeft      =   2280
   ClientTop       =   2580
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   12585
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7515
      Left            =   720
      OleObjectBlob   =   "frmReport1.frx":0000
      TabIndex        =   0
      Top             =   450
      Width           =   13260
   End
End
Attribute VB_Name = "frmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim t1 As Date
Dim i As Integer
Dim avgtime As Double
Dim dtime As Integer
Dim totTime As Integer
Dim lTime As Integer
Dim sTime As Integer
Dim X(10 To 50, 1 To 2)
For i = 10 To 50
    X(i, 1) = CStr(i) & "sec"
    X(i, 2) = 0
Next i
lTime = 0
sTime = 10000
Set rs1 = db.OpenRecordset("select Event,EventTime from StartStop where eventTime > #3/5/2004# order by EventTime")
t1 = rs1("EventTime")
rs1.MoveNext
Do Until rs1.EOF
    If rs1("Event") = "STOP" Then
        dtime = DateDiff("s", t1, rs1("EventTime"))
        X(dtime, 2) = X(dtime, 2) + 1
        If dtime > lTime Then lTime = dtime
        If dtime < sTime Then sTime = dtime
        totTime = totTime + dtime
        i = i + 1
    Else
        t1 = rs1("EventTime")
    End If
    rs1.MoveNext
Loop
avgtime = totTime / i
MSChart1.ChartData = X
End Sub


