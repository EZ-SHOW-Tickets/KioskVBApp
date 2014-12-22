VERSION 5.00
Object = "{7C6896B9-38B9-11D2-91BB-00C0A8356885}#1.0#0"; "DSINTX.OCX"
Begin VB.Form frmPIN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   3195
   ClientTop       =   4005
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPIN.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   165
      Top             =   195
   End
   Begin DSINTXLib.Dsintx Dsintx1 
      Left            =   180
      Top             =   990
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1323
      _StockProps     =   0
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please follow instructions on the PIN Pad screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3495
      TabIndex        =   0
      Top             =   3105
      Width           =   8235
   End
End
Attribute VB_Name = "frmPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'If CardType <> "120" Then
    frmBankAuthorization.Show
    Unload Me
'Else
''GET PIN
'End If

End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
frmPrinting.Show
Unload Me
End Sub


