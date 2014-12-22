VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmSetUpNet 
   Caption         =   "Set Up Network"
   ClientHeight    =   6090
   ClientLeft      =   4035
   ClientTop       =   3135
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8940
   Begin Threed.SSCommand cmdCancel 
      Height          =   495
      Left            =   6615
      TabIndex        =   0
      Top             =   5460
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   873
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CANCEL"
   End
   Begin VB.Label lblTitles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "WorkStation Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   375
      TabIndex        =   4
      Top             =   315
      Width           =   2355
   End
   Begin VB.Label lblTitles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   750
      Width           =   2355
   End
   Begin VB.Label lblComputername 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2805
      TabIndex        =   2
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label lblServerName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2805
      TabIndex        =   1
      Top             =   735
      Width           =   3000
   End
End
Attribute VB_Name = "frmSetUpNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
