VERSION 5.00
Object = "{7C6896B9-38B9-11D2-91BB-00C0A8356885}#1.0#0"; "DSINTX.OCX"
Begin VB.Form frmPINPAD 
   Caption         =   "PIN PAD"
   ClientHeight    =   1725
   ClientLeft      =   2145
   ClientTop       =   4605
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3225
   Begin DSINTXLib.Dsintx Dsintx1 
      Left            =   180
      Top             =   120
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   767
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPINPAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim x As Integer

Dsintx1.ReleaseConn
For i = 1 To 1000
    x = DoEvents
Next i
End Sub


