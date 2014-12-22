VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmSpecifyTickets 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10815
   ClientLeft      =   3240
   ClientTop       =   2370
   ClientWidth     =   15210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSpecifyTickets.frx":0000
   ScaleHeight     =   10815
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer EMail_Timer 
      Left            =   120
      Top             =   1800
   End
   Begin VB.Timer ShowTimer 
      Left            =   120
      Top             =   1200
   End
   Begin VB.Timer Zip_Timer 
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer_Verify 
      Left            =   105
      Top             =   120
   End
   Begin VB.Frame FrameTicketSum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   4
      Left            =   7605
      TabIndex        =   36
      Top             =   7545
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   4
         Left            =   4290
         TabIndex        =   37
         Top             =   30
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.Label lblCategoryAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "3 @ $502.00 = $536.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   150
         TabIndex        =   38
         Top             =   135
         Width           =   4125
      End
   End
   Begin VB.Frame FrameTicketSum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   3
      Left            =   7605
      TabIndex        =   31
      Top             =   5940
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   3
         Left            =   4290
         TabIndex        =   32
         Top             =   30
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.Label lblCategoryAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "3 @ $12.00 = $36.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   150
         TabIndex        =   33
         Top             =   135
         Width           =   4125
      End
   End
   Begin Threed.SSCommand cmdCancelTrans 
      Height          =   1080
      Left            =   915
      TabIndex        =   30
      Top             =   8700
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   1905
      _Version        =   196609
      BackColor       =   8421631
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
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   0
      Left            =   3840
      TabIndex        =   24
      Top             =   690
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2646
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adult $15.00"
      BevelWidth      =   2
      BorderWidth     =   5
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label lblTax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$13.78 + $1.22 tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   41
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label lblSubButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age 5 - 12 years old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   25
         Top             =   1080
         Width           =   3390
      End
   End
   Begin Threed.SSPanel PanelKeyPad 
      Height          =   2250
      Left            =   7560
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3969
      _Version        =   196609
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdKeyPadCancel 
         Height          =   840
         Left            =   4905
         TabIndex        =   23
         Top             =   1260
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1482
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   8
         Left            =   3735
         TabIndex        =   10
         Top             =   1268
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   7
         Left            =   2550
         TabIndex        =   9
         Top             =   1268
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   6
         Left            =   1380
         TabIndex        =   8
         Top             =   1268
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   5
         Left            =   210
         TabIndex        =   7
         Top             =   1268
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   4
         Left            =   4890
         TabIndex        =   6
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   3
         Left            =   3720
         TabIndex        =   5
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   2
         Left            =   2550
         TabIndex        =   4
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   847
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label LblPickNumber 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select number of Senior tickets (0 to cancel)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   5355
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   1
      Left            =   7575
      TabIndex        =   20
      Top             =   2708
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   1
         Left            =   4305
         TabIndex        =   21
         Top             =   30
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.Label lblCategoryAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "3 @ $12.00 = $36.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   225
         TabIndex        =   22
         Top             =   90
         Width           =   4035
      End
   End
   Begin VB.Frame FrameTicketSum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   2
      Left            =   7575
      TabIndex        =   16
      Top             =   4328
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   2
         Left            =   4290
         TabIndex        =   17
         Top             =   30
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.Label lblCategoryAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "3 @ $12.00 = $36.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   135
         Width           =   4125
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   0
      Left            =   7605
      TabIndex        =   14
      Top             =   1095
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   0
         Left            =   4275
         TabIndex        =   15
         Top             =   30
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin VB.Label lblCategoryAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "3 @ $12.00 = $36.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   105
         Width           =   4035
      End
   End
   Begin Threed.SSCommand cmdPurchase 
      Height          =   1725
      Left            =   7740
      TabIndex        =   12
      Top             =   8760
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3043
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   65535
      BackColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PURCHASE THESE TICKETS"
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   1
      Left            =   3855
      TabIndex        =   26
      Top             =   2310
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2646
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adult $15.00"
      BevelWidth      =   2
      BorderWidth     =   5
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label lblTax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$4.59 + $0.41 tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   42
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label lblSubButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age 5 - 12 years old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   27
         Top             =   1080
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   2
      Left            =   3855
      TabIndex        =   28
      Top             =   3930
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2646
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adult $15.00"
      BevelWidth      =   2
      BorderWidth     =   5
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label lblTax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$4.49 + $0.41 tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   43
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label lblSubButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age 5 - 12 years old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   29
         Top             =   1080
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   3
      Left            =   3885
      TabIndex        =   34
      Top             =   5535
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2646
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adult $15.00"
      BevelWidth      =   2
      BorderWidth     =   5
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label lblSubButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age 5 - 12 years old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   35
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   4
      Left            =   3885
      TabIndex        =   39
      Top             =   7140
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2646
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adult $15.00"
      BevelWidth      =   2
      BorderWidth     =   5
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label lblSubButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Age 5 - 12 years old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   105
         TabIndex        =   40
         Top             =   1005
         Width           =   3390
      End
   End
   Begin VB.Image imgProgram 
      Height          =   3570
      Left            =   600
      Picture         =   "frmSpecifyTickets.frx":A06F2
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label LblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Purchase Total: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8415
      TabIndex        =   11
      Top             =   8460
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a ticket category that you would like to purchase"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   675
      TabIndex        =   0
      Top             =   90
      Width           =   13125
   End
End
Attribute VB_Name = "frmSpecifyTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click(Index As Integer)
    Timer_Verify.Enabled = False
    TotalTicketSold(Index) = 0
    FrameTicketSum(Index).Visible = False
    TotalPrice = 0
    For i = 0 To UBound(TicketTypes, 2)
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        panelTicketButton(i).Enabled = True
    Next i
    LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
    If TotalPrice > 0 Then
        cmdPurchase.Visible = True
        LblTotal.Visible = True
        lblInstructions.Caption = "Please select another category or press PURCHASE THESE TICKETS"
    Else
        LblTotal.Visible = False
        cmdPurchase.Visible = False
        lblInstructions.Caption = "Please select a ticket category that you would like to purchase"
   
    End If
    Call LogClick("SpecifyTickets", "Cancel")
    Timer_Verify.Enabled = True

End Sub

Private Sub cmdCancelTrans_Click()
Call LogClick("SpecifyTickets", "CancelTransaction")
'' CLEAR ALL TICKETS
Erase TicketTypes
Erase TotalTicketSold
TotTicketCount = 0
TicketCount = 0
bChange = False
'''''''''''''''''''
Timer_Verify.interval = 0
Timer_Verify.Enabled = False
If bDebug Then MsgBox "Back to WELCOME"
'If bMultiple Then
'    frmWelcomeMult.Show
'Else
    'If UseAOP Then
        frmFirstPageAOP.Show
    'Else
    '    frmFirstPage.Show
    'End If
'End If
Unload Me
End Sub


Private Sub cmdKeyPad_Click(Index As Integer)
    Dim TotalPrice As Currency
    Timer_Verify.Enabled = False
    PanelKeyPad.Visible = False
    LblPickNumber.Visible = False
    
    For i = 0 To FrameTicketSum.Count - 1
        FrameTicketSum(i).Enabled = True
    Next i

 '   If Index < 9 Then
        lblCategoryAmount(SelectedTicket).Caption = CStr(Index + 1) & " tickets @" & Format(TicketTypes(3, SelectedTicket), "$#.00") & "=" & Format((Index + 1) * TicketTypes(3, SelectedTicket), "$#.00")
        FrameTicketSum(SelectedTicket).Visible = True
        TotalTicketSold(SelectedTicket) = Index + 1
'    Else
'         TotalTicketSold(SelectedTicket) = 0
'         FrameTicketSum(SelectedTicket).Visible = False
'    End If
    'cmdTicketBtn(SelectedTicket).BackColor = &HC0C0C0
    panelTicketButton(SelectedTicket).BackColor = &HC0C0C0
    If TicketTypes(2, SelectedTicket) = "Program" Then panelTicketButton(SelectedTicket).ForeColor = &HFF&

    For i = 0 To UBound(TicketTypes, 2)
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        'cmdTicketBtn(i).Enabled = True
        panelTicketButton(i).Enabled = True
        lblSubButton(i).Enabled = True

    Next i
    LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
    If TotalPrice > 0 Then
        cmdPurchase.Visible = True
        cmdPurchase.Enabled = True
        LblTotal.Visible = True
        lblInstructions.Caption = "Please select another category or press PURCHASE THESE TICKETS"
    Else
        LblTotal.Visible = False
        cmdPurchase.Visible = False
        lblInstructions.Caption = "Please select a ticket category that you would like to purchase"
    
    End If
   Call LogClick("SpecifyTickets", cmdKeyPad(Index).Caption)
    Timer_Verify.Enabled = True
End Sub

Private Sub cmdKeyPadCancel_Click()
    Dim TotalPrice As Currency
    Timer_Verify.Enabled = False
    For i = 0 To FrameTicketSum.Count - 1
        FrameTicketSum(i).Enabled = True
    Next i
    
    PanelKeyPad.Visible = False
    LblPickNumber.Visible = False

    TotalTicketSold(SelectedTicket) = 0
    FrameTicketSum(SelectedTicket).Visible = False
    'cmdTicketBtn(SelectedTicket).BackColor = &HC0C0C0
    panelTicketButton(SelectedTicket).BackColor = &HC0C0C0
    For i = 0 To UBound(TicketTypes, 2)
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        'cmdTicketBtn(i).Enabled = True
        panelTicketButton(i).Enabled = True
        lblSubButton(i).Enabled = True

    Next i
    LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
    If TotalPrice > 0 Then
        cmdPurchase.Visible = True
        cmdPurchase.Enabled = True
        LblTotal.Visible = True
        lblInstructions.Caption = "Please select another category or press PURCHASE THESE TICKETS"
    Else
        LblTotal.Visible = False
        cmdPurchase.Visible = False
        lblInstructions.Caption = "Please select a ticket category that you would like to purchase"
    
    End If
   Call LogClick("SpecifyTickets", "Cancel")
   Timer_Verify.Enabled = False

End Sub


Private Sub cmdPurchase_Click()
Timer_Verify.Enabled = False
Call LogClick("SpecifyTickets", "Purchase")
For i = 0 To UBound(TicketTypes, 2)
    If TotalTicketSold(i) > 0 Then
        ''For COUPONS
        If TicketTypes(1, i) = 26 Then
            frmCouponCode.Show 1
            If Len(CouponCode) = 1 Then
             Unload Me
             frmSpecifyTickets.Show
             Exit Sub
            End If
        End If
    End If
Next i
frmVerify.Show
Unload Me
End Sub


Private Sub EMail_Timer_Timer()
    EMail_Timer.interval = 0
    EMail_Timer.Enabled = False
    frmemail.Show 1
    If bRestart Then
        Timer_Verify.interval = 100
        bRestart = False
    Else
        Timer_Verify.interval = 20000
    End If
    Timer_Verify.Enabled = True

End Sub

Private Sub Form_Load()
Dim i
Dim Index As Integer
Dim BottomButton As Integer
Dim ShowFirst As Boolean
Dim BaseTicket As Double
ShowFirst = False
CouponCode = "0"
If bDebug Then MsgBox "Form Load"
Call LogClick("SpecifyTickets", "Load")
bRestart = False
'
'Ask for ZIP CODE if set
'
If GetZip Then
    Zip_Timer.interval = 10
    Zip_Timer.Enabled = True
Else
    Zip_Timer.Enabled = False
    If bGetEMail = False Then
        Timer_Verify.interval = 15000
        Timer_Verify.Enabled = True
    End If
End If
If bGetEMail Then
    EMail_Timer.interval = 10
    EMail_Timer.Enabled = True
Else
    EMail_Timer.Enabled = False
    Timer_Verify.interval = 15000
    Timer_Verify.Enabled = True
End If

If bDebug Then MsgBox "Getting Ticket info"
'
'Get TICKET CATEGORIES for this show/time
'
'' DO NOT LOAD PROGRAM TICKET
'Set rsTickets = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategoryID,TicketCategory,TicketPrice,subButtonText,CategoryName from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and #" & Now() & "# <= ShowDateStop and #" & Now() & "# >= ShowDateStart and TICKET_CATEGORIES.TicketCategoryID < 10 order by SHOW_TICKET_PRICES.ListSequence")
Set rsTickets = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategoryID,TicketCategory,TicketPrice,subButtonText,CategoryName from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and #" & Now() & "# <= ShowDateStop and #" & Now() & "# >= ShowDateStart order by SHOW_TICKET_PRICES.ListSequence")
'
' Populate ticket arrays
'
i = -1
Do Until rsTickets.EOF
    i = i + 1
    ReDim Preserve TicketTypes(5, i)
    TicketTypes(1, i) = rsTickets("TicketCategoryID")
    TicketTypes(2, i) = rsTickets("TicketCategory")
    TicketTypes(3, i) = rsTickets("TicketPrice")
    If Not IsNull(rsTickets("subbuttontext")) Then
        TicketTypes(4, i) = rsTickets("subbuttontext")
    Else
        TicketTypes(4, i) = ""
    End If
    TicketTypes(5, i) = rsTickets("CategoryName")
    
    rsTickets.MoveNext
Loop
'
' Populate Buttons on screen
'
For i = 0 To UBound(TicketTypes, 2)
    lblSubButton(i) = TicketTypes(4, i)
    panelTicketButton(i).Caption = TicketTypes(2, i) & " " & Format(TicketTypes(3, i), "$#.00")
    panelTicketButton(i).Visible = True
'    If TicketTypes(2, i) = "Program" Then
'        panelTicketButton(i).ForeColor = &HFF&
'        imgProgram.Visible = True
'    End If
    If EventCity = "NEW YORK" Then
        BaseTicket = CDec(TicketTypes(3, i)) / (1# + 0.08875)
        lblTax(i) = Format(BaseTicket, "$#.00") & " plus " & Format(CDec(TicketTypes(3, i)) - BaseTicket, "$#.00") & " tax"
        'lblTax(i) = Format(CDec(TicketTypes(3, i)) / (1# + 0.08875), "$#.00") & " plus " & Format(CDec(TicketTypes(3, i)) * 0.08875, "$#.00") & " tax"
        lblTax(i).Visible = True
        imgProgram.Picture = LoadPicture(App.path & "/NYProgram.jpg")
        imgProgram.Visible = True
    Else
        lblTax(i).Visible = False
        imgProgram.Visible = False
    End If
    BottomButton = i
Next i
'
'' bChange - Set in Verify screen
If bChange Then
'' Put up last selection and allow change
    TotalPrice = 0
    For i = 0 To UBound(TotalTicketSold, 1)
        If TotalTicketSold(i) > 0 Then
            lblCategoryAmount(i).Caption = TotalTicketSold(i) & " tickets @" & Format(TicketTypes(3, i), "$#.00") & "=" & Format(TotalTicketSold(i) * TicketTypes(3, i), "$#.00")
            FrameTicketSum(i).Visible = True
            TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
            panelTicketButton(i).Enabled = True
        End If
    Next i
    LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
    If TotalPrice > 0 Then
        cmdPurchase.Visible = True
        LblTotal.Visible = True
    Else
        LblTotal.Visible = False
        cmdPurchase.Visible = False
    
    End If
    'Turn off change mode
     bChange = False
Else
  ' Not in Change mode
    TotalTicketPrice = 0#
    ReDim TotalTicketSold(UBound(TicketTypes, 2))
    For i = 0 To UBound(TicketTypes, 2)
        TotalTicketSold(i) = 0
    Next i
    For i = 0 To FrameTicketSum.Count - 1
        FrameTicketSum(i).Enabled = False
    Next i
    Index = 0
    SelectedTicket = Index
    For i = 0 To panelTicketButton.Count - 1
                'Put up first number pad
                panelTicketButton(i).BackColor = &HC0FFC0
                panelTicketButton(i).Enabled = True
                If ShowFirst Then
                    If i = 0 Then
                        PanelKeyPad.Top = panelTicketButton(i).Top
                        PanelKeyPad.Visible = True
                        LblPickNumber.Caption = "Select number of " & TicketTypes(2, i) & " tickets (cancel for 0)"
                        FrameTicketSum(SelectedTicket).Visible = False
                        LblPickNumber.Visible = True
                        ShowTimer.interval = 4000
                        ShowTimer.Enabled = True
                    Else
                        panelTicketButton(i).Enabled = False
                        lblSubButton(i).Enabled = False
                    End If
        
                 End If
    Next i


End If
''''' TOP OF PURCHASE BUTTON
cmdPurchase.Top = panelTicketButton(BottomButton).Top + panelTicketButton(BottomButton).Height + 300
LblTotal.Top = panelTicketButton(BottomButton).Top + panelTicketButton(BottomButton).Height

Call LogClick("SpecifyTickets", "Wait For Select")

'''''''''''''''''''
' FIRST TIME
'Was set to FALSE but it shouldn't be
'Timer_Verify.Enabled = False
''''''''''''''''''''
If bGetEMail = False Then
    Timer_Verify.Enabled = True
    Timer_Verify.interval = 15000
End If
''''''''''''''''''''''''
'
' Put up first button select pad
'
'Index = 0
'SelectedTicket = Index
'TotalTicketSold(Index) = 0
'FrameTicketSum(Index).Visible = False
'
'turn off SOLD Frame
'
'For i = 0 To FrameTicketSum.Count - 1
'    FrameTicketSum(i).Enabled = False
'Next i

'TotalPrice = 0
'For i = 0 To UBound(TicketTypes, 2) - 1
'        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
'        panelTicketButton(i).Enabled = True
'Next i
'LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
'If TotalPrice > 0 Then
'    cmdPurchase.Visible = True
'    LblTotal.Visible = True
'    lblInstructions.Caption = "Please select another category or press PURCHASE THESE TICKETS"
'Else
'    LblTotal.Visible = False
'    cmdPurchase.Visible = False
'    lblInstructions.Caption = "Please select a ticket category that you would like to purchase"
'End If

'If Not bChange Then
'        For i = 0 To panelTicketButton.Count - 1
'            If i = Index Then
'                'Put up first number pad
'                panelTicketButton(i).BackColor = &HC0FFC0
'                panelTicketButton(i).Enabled = True
'                PanelKeyPad.Top = panelTicketButton(Index).Top
'                PanelKeyPad.Visible = True
'                LblPickNumber.Caption = "Select number of " & TicketTypes(2, SelectedTicket) & " tickets (0 to cancel)"
'                FrameTicketSum(SelectedTicket).Visible = False
'                LblPickNumber.Visible = True
'           Else
'                panelTicketButton(i).Enabled = False
'                lblSubButton(i).Enabled = False
'
'            End If
'        Next i
'End If


'''''''''''''''''''
End Sub






Private Sub lblSubButton_Click(Index As Integer)
    panelTicketButton_Click (Index)
End Sub

Private Sub lblTax_Click(Index As Integer)
panelTicketButton_Click (Index)
End Sub

Private Sub panelTicketButton_Click(Index As Integer)
Timer_Verify.Enabled = False
SelectedTicket = Index
TotalTicketSold(Index) = 0
FrameTicketSum(Index).Visible = False
For i = 0 To FrameTicketSum.Count - 1
    FrameTicketSum(i).Enabled = False
Next i
TotalPrice = 0
For i = 0 To UBound(TicketTypes, 2) - 1
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        panelTicketButton(i).Enabled = True
Next i
LblTotal.Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
If TotalPrice > 0 Then
    cmdPurchase.Visible = True
    LblTotal.Visible = True
    lblInstructions.Caption = "Please select another category or press PURCHASE THESE TICKETS"
Else
    LblTotal.Visible = False
    cmdPurchase.Visible = False
    lblInstructions.Caption = "Please select a ticket category that you would like to purchase"
End If

For i = 0 To panelTicketButton.Count - 1
    If i = Index Then
        panelTicketButton(i).BackColor = &HC0FFC0
        panelTicketButton(i).Enabled = True
        PanelKeyPad.Top = panelTicketButton(Index).Top
'        If InStr(panelTicketButton(Index).Caption, "Family") > 0 Then
'            For j = 2 To 8
'                PanelKeyPad(j).Enabled = False
'            Next j
'        Else
'            For j = 2 To 8
'                PanelKeyPad(j).Enabled = True
'            Next j
'        End If
        PanelKeyPad.Visible = True
        LblPickNumber.Caption = "Select number of " & TicketTypes(2, SelectedTicket) & " tickets (cancel for 0)"
        FrameTicketSum(SelectedTicket).Visible = False
        LblPickNumber.Visible = True
   Else
        panelTicketButton(i).Enabled = False
        lblSubButton(i).Enabled = False
    End If
Next i
cmdPurchase.Enabled = False
Call LogClick("SpecifyTickets", panelTicketButton(Index).Caption)
Timer_Verify.Enabled = True
End Sub


Private Sub ShowTimer_Timer()
    ShowTimer.interval = 0
    ShowTimer.Enabled = False

    For i = 0 To panelTicketButton.Count - 1
                'Put up first number pad
                panelTicketButton(i).BackColor = &HC0FFC0
                panelTicketButton(i).Enabled = True
                PanelKeyPad.Visible = False

    Next i

End Sub

Private Sub Timer_Verify_Timer()
Timer_Verify.interval = 0
Timer_Verify.Enabled = False
If bDebug Then MsgBox "Time Out in Verify"
bChange = False
cmdCancelTrans_Click

End Sub


Private Sub Zip_Timer_Timer()
    Zip_Timer.interval = 0
    Zip_Timer.Enabled = False
    frmZipCode.Show 1
    
    If bRestart Then
        Timer_Verify.interval = 100
        bRestart = False
    Else
        Timer_Verify.interval = 20000
    End If
    Timer_Verify.Enabled = True

End Sub


