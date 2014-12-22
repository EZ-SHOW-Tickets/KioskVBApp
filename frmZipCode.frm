VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmZipCode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Zip Code"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimeOutTimer 
      Left            =   240
      Top             =   1425
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   915
      Left            =   4680
      TabIndex        =   10
      Top             =   5100
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1614
      _Version        =   196609
      Font3D          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLEAR"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   0
      Left            =   5475
      TabIndex        =   0
      Top             =   3825
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   2565
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   2565
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   3
      Left            =   2835
      TabIndex        =   3
      Top             =   2565
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   4
      Left            =   4140
      TabIndex        =   4
      Top             =   2565
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   5
      Left            =   5445
      TabIndex        =   5
      Top             =   2565
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3825
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   7
      Left            =   1545
      TabIndex        =   7
      Top             =   3825
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "7"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   8
      Left            =   2850
      TabIndex        =   8
      Top             =   3825
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "8"
   End
   Begin Threed.SSCommand cmdZip 
      Height          =   1140
      Index           =   9
      Left            =   4155
      TabIndex        =   9
      Top             =   3825
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "9"
   End
   Begin Threed.SSCommand cmdEnter 
      Height          =   1320
      Left            =   300
      TabIndex        =   11
      Top             =   5130
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2328
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   49152
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SUBMIT"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   570
      Left            =   4170
      TabIndex        =   15
      Top             =   795
      Width           =   1860
   End
   Begin VB.Label lblInstructions2 
      BackStyle       =   0  'Transparent
      Caption         =   "and then press"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   660
      TabIndex        =   14
      Top             =   765
      Width           =   3465
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER ZIP CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   60
      TabIndex        =   13
      Top             =   15
      Width           =   7140
   End
   Begin VB.Label lblZipCode 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2160
      TabIndex        =   12
      Top             =   1470
      Width           =   2265
   End
End
Attribute VB_Name = "frmZipCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdClear_Click()

lblZipCode.Caption = ""
cmdEnter.Enabled = False
End Sub

Private Sub cmdEnter_Click()
TimeOutTimer.Enabled = False
ZipCode = lblZipCode.Caption
Unload Me
End Sub

Private Sub cmdZip_Click(Index As Integer)
If Len(Trim(lblZipCode.Caption)) < 5 Then
    'TimeOutTimer.interval = 12000
    lblZipCode.Caption = lblZipCode.Caption & CStr(Index)
    If Len(Trim(lblZipCode.Caption)) = 5 Then
        cmdEnter.Enabled = True
    End If
End If
End Sub


Private Sub Form_Load()
Refresh

TimeOutTimer.interval = 15000

End Sub


Private Sub TimeOutTimer_Timer()
TimeOutTimer.Enabled = False
ZipCode = ""
bRestart = True
Unload Me

End Sub


