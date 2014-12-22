VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmSecurityTix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security Tickets"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSecurity 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   1
      Text            =   "6"
      Top             =   360
      Width           =   645
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "DONE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin Threed.SSCommand cmdSecurity 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "+"
   End
   Begin VB.Label lblSecurity 
      Alignment       =   1  'Right Justify
      Caption         =   "Security Cutoff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1950
   End
End
Attribute VB_Name = "frmSecurityTix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
iMaxTix = Int(txtSecurity.text)

Unload Me
End Sub

Private Sub cmdSecurity_Click()
If Int(txtSecurity.text) = "9" Then
   txtSecurity.text = "1"
Else
   txtSecurity.text = Trim(str(Int(txtSecurity.text) + 1))
   
End If

End Sub


