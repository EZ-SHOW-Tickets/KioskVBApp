VERSION 5.00
Object = "{7C6896B9-38B9-11D2-91BB-00C0A8356885}#1.0#0"; "DSINTX.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   9825
   ClientTop       =   3705
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4080
   Begin VB.Timer Timer1 
      Left            =   765
      Top             =   660
   End
   Begin DSINTXLib.Dsintx Dsintx1 
      Left            =   2280
      Top             =   600
      _Version        =   65536
      _ExtentX        =   1217
      _ExtentY        =   1191
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Interval = 1000
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    Timer1.Enabled = False
    Dsintx1.SetTfield "PinPadPort", 4
    Dsintx1.InitTransaction
    Dsintx1.SetTfield "ReadCard", "y"
    xreturn = Dsintx1.ProcessTransaction
    AccountNumber = Dsintx1.GetTfield("Track2_Acc")
    db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
    Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
    SessionID = rs(0)
    Call LogClick("Welcome", "SwipeCard")
    frmSpecifyTickets.Show
    Unload Me
End Sub


