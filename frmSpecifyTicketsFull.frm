VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmSpecifyTicketsFull 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9945
   ClientLeft      =   1065
   ClientTop       =   975
   ClientWidth     =   13545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSpecifyTicketsFull.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   0
      Left            =   3855
      TabIndex        =   34
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
         TabIndex        =   35
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel PanelKeyPad 
      Height          =   1995
      Left            =   7560
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   3519
      _Version        =   196609
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdKeyPadCancel 
         Height          =   720
         Left            =   3570
         TabIndex        =   33
         Top             =   1110
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
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
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   8
         Left            =   2730
         TabIndex        =   10
         Top             =   1125
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   7
         Left            =   1890
         TabIndex        =   9
         Top             =   1125
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   6
         Left            =   1050
         TabIndex        =   8
         Top             =   1125
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   5
         Left            =   210
         TabIndex        =   7
         Top             =   1125
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   4
         Left            =   3570
         TabIndex        =   6
         Top             =   420
         Width           =   825
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   3
         Left            =   2730
         TabIndex        =   5
         Top             =   420
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   2
         Left            =   1890
         TabIndex        =   4
         Top             =   420
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   1
         Left            =   1050
         TabIndex        =   3
         Top             =   420
         Width           =   840
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   840
      End
      Begin VB.Label LblPickNumber 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select number of Senior tickets (0 to cancel)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Width           =   4305
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   5
      Left            =   7560
      TabIndex        =   30
      Top             =   9195
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   5
         Left            =   4290
         TabIndex        =   31
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
         Index           =   5
         Left            =   150
         TabIndex        =   32
         Top             =   135
         Width           =   4110
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   4
      Left            =   7590
      TabIndex        =   27
      Top             =   7568
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   4
         Left            =   4290
         TabIndex        =   28
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
         Index           =   4
         Left            =   150
         TabIndex        =   29
         Top             =   135
         Width           =   4110
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   1
      Left            =   7575
      TabIndex        =   24
      Top             =   2708
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   1
         Left            =   4305
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   90
         Width           =   4035
      End
   End
   Begin Threed.SSCommand cmdCancelTrans 
      Height          =   930
      Left            =   210
      TabIndex        =   23
      Top             =   10335
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   1640
      _Version        =   196609
      Font3D          =   2
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CANCEL"
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   3
      Left            =   7560
      TabIndex        =   19
      Top             =   5948
      Visible         =   0   'False
      Width           =   5805
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Index           =   3
         Left            =   4290
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   135
         Width           =   4110
      End
   End
   Begin VB.Frame FrameTicketSum 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      BackColor       =   &H00E0E0E0&
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
         TabIndex        =   22
         Top             =   105
         Width           =   4035
      End
   End
   Begin Threed.SSCommand cmdPurchase 
      Height          =   1020
      Left            =   8415
      TabIndex        =   12
      Top             =   10335
      Visible         =   0   'False
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   1799
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      TabIndex        =   36
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
         TabIndex        =   37
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   2
      Left            =   3855
      TabIndex        =   38
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
         TabIndex        =   39
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   3
      Left            =   3855
      TabIndex        =   40
      Top             =   5565
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
         TabIndex        =   41
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   4
      Left            =   3855
      TabIndex        =   42
      Top             =   7185
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
         TabIndex        =   43
         Top             =   1005
         Width           =   3390
      End
   End
   Begin Threed.SSPanel panelTicketButton 
      Height          =   1500
      Index           =   5
      Left            =   3855
      TabIndex        =   44
      Top             =   8805
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
         Index           =   5
         Left            =   105
         TabIndex        =   45
         Top             =   1005
         Width           =   3390
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   -30
      Top             =   0
      Width           =   2685
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
      Top             =   10035
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
Attribute VB_Name = "frmSpecifyTicketsFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click(Index As Integer)
    TotalTicketSold(Index) = 0
    FrameTicketSum(Index).Visible = False

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
    Call LogClick("SpecifyTickets", "Cancel")

End Sub

Private Sub cmdCancelTrans_Click()
Call LogClick("SpecifyTickets", "CancelTransaction")

frmWelcome.Show
Unload Me
End Sub


Private Sub cmdKeyPad_Click(Index As Integer)
    Dim TotalPrice As Currency
    
    PanelKeyPad.Visible = False
    LblPickNumber.Visible = False

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
    For i = 0 To UBound(TicketTypes, 2)
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        'cmdTicketBtn(i).Enabled = True
        panelTicketButton(i).Enabled = True
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
 
End Sub

Private Sub cmdKeyPadCancel_Click()
    Dim TotalPrice As Currency
    
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

End Sub


Private Sub cmdPurchase_Click()
Call LogClick("SpecifyTickets", "Purchase")

frmVerify.Show
Unload Me
End Sub


Private Sub Form_Load()
Dim i

Set rsTickets = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategoryID,TicketCategory,TicketPrice,subButtonText from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and (#" & Now() & "# between ShowDateStart and ShowDateStop) order by SHOW_TICKET_PRICES.ListSequence")
If rsTickets.EOF Then
  Set rsTickets = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategoryID,TicketCategory,TicketPrice,subButtonText from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & "7" & " order by SHOW_TICKET_PRICES.ListSequence")
End If
i = -1
Do Until rsTickets.EOF
    i = i + 1
    ReDim Preserve TicketTypes(4, i)
    TicketTypes(1, i) = rsTickets("TicketCategoryID")
    TicketTypes(2, i) = rsTickets("TicketCategory")
    TicketTypes(3, i) = rsTickets("TicketPrice")
    If Not IsNull(rsTickets("subbuttontext")) Then
        TicketTypes(4, i) = rsTickets("subbuttontext")
    Else
        TicketTypes(4, i) = ""
    End If
    rsTickets.MoveNext
Loop

For i = 0 To UBound(TicketTypes, 2)
    'cmdTicketBtn(i).Caption = TicketTypes(2, i) & " " & Format(TicketTypes(3, i), "$#.00")
    'cmdTicketBtn(i).Visible = True
    lblSubButton(i) = TicketTypes(4, i)
    panelTicketButton(i).Caption = TicketTypes(2, i) & " " & Format(TicketTypes(3, i), "$#.00")
    panelTicketButton(i).Visible = True
Next i
If bChange Then
    For i = 0 To UBound(TotalTicketSold, 1)
        If TotalTicketSold(i) > 0 Then
            lblCategoryAmount(i).Caption = TotalTicketSold(i) & " tickets @" & Format(TicketTypes(3, i), "$#.00") & "=" & Format(TotalTicketSold(i) * TicketTypes(3, i), "$#.00")
            FrameTicketSum(i).Visible = True
            TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
            'cmdTicketBtn(i).Enabled = True
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
Else
    TotalTicketPrice = 0#
    ReDim TotalTicketSold(UBound(TicketTypes, 2))
    For i = 0 To UBound(TicketTypes, 2)
        TotalTicketSold(i) = 0
    Next i
    bChange = False
End If
End Sub


Private Sub LabelEnd_DblClick()
End
End Sub


Private Sub Image1_DblClick()
End
End Sub


Private Sub lblSubButton_Click(Index As Integer)
    panelTicketButton_Click (Index)
End Sub

Private Sub panelTicketButton_Click(Index As Integer)
SelectedTicket = Index
TotalTicketSold(Index) = 0
FrameTicketSum(Index).Visible = False
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
        PanelKeyPad.Visible = True
        LblPickNumber.Caption = "Select number of " & TicketTypes(2, SelectedTicket) & " tickets (0 to cancel)"
        FrameTicketSum(SelectedTicket).Visible = False
        LblPickNumber.Visible = True
   Else
        panelTicketButton(i).Enabled = False
    End If
Next i
cmdPurchase.Enabled = False
Call LogClick("SpecifyTickets", panelTicketButton(Index).Caption)

End Sub


