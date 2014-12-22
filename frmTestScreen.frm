VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmTestScreen 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   465
   ClientTop       =   1725
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   13545
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand SSCommand1 
      Height          =   1830
      Left            =   1080
      TabIndex        =   0
      Top             =   4140
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   3228
      _Version        =   196609
      BackColor       =   255
      Caption         =   "BACK"
   End
End
Attribute VB_Name = "frmTestScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()
frmTestInterface.Show
Unload Me
End Sub


