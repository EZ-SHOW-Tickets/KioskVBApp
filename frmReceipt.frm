VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmReceipt 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Receipt"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   255
   End
   Begin Threed.SSCommand cmdYesOrNo 
      Height          =   1305
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   1515
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2302
      _Version        =   196609
      Font3D          =   4
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "YES"
   End
   Begin Threed.SSCommand cmdYesOrNo 
      Height          =   1305
      Index           =   1
      Left            =   4365
      TabIndex        =   1
      Top             =   1515
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2302
      _Version        =   196609
      Font3D          =   4
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NO"
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DO YOU WANT A PRINTED RECEIPT WITH YOUR TICKETS?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1020
      Left            =   555
      TabIndex        =   2
      Top             =   285
      Width           =   7590
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click(Index As Integer)

End Sub


Private Sub cmdYesOrNo_Click(Index As Integer)
If Index = 0 Then
  bPrintReceipt = True
Else
  bPrintReceipt = False
End If
Unload Me
End Sub


Private Sub Form_Load()

Timer1.Interval = 5000
Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()
Call cmdYesOrNo_Click(1)
End Sub


