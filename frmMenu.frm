VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMenu 
   Caption         =   "WalkUp Systems"
   ClientHeight    =   9135
   ClientLeft      =   1875
   ClientTop       =   1395
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkZipCode 
      Caption         =   "Get ZIP"
      Height          =   345
      Left            =   7650
      TabIndex        =   59
      Top             =   4140
      Width           =   1575
   End
   Begin VB.CommandButton cmdZip 
      Caption         =   "Zip"
      Height          =   570
      Left            =   9675
      TabIndex        =   58
      Top             =   8490
      Width           =   1050
   End
   Begin VB.TextBox txt_xxxx 
      Height          =   465
      Left            =   8805
      TabIndex        =   57
      Text            =   "1"
      Top             =   8640
      Width           =   585
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   45
      Top             =   4365
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Threed.SSCommand cmdBatch 
      Height          =   615
      Left            =   1785
      TabIndex        =   55
      Top             =   8355
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "Batch"
   End
   Begin VB.CheckBox AN 
      Caption         =   "AN"
      Height          =   345
      Left            =   6615
      TabIndex        =   54
      Top             =   4155
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CheckBox CheckDiscover 
      Caption         =   "Allow Discover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   495
      TabIndex        =   52
      Top             =   4455
      Width           =   2310
   End
   Begin VB.CheckBox CheckVISAMC 
      Caption         =   "VISA/MC ONLY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2925
      TabIndex        =   51
      Top             =   4515
      Width           =   2085
   End
   Begin VB.CheckBox checkDEMO 
      Caption         =   "DEMO ONLY"
      Height          =   345
      Left            =   5280
      TabIndex        =   50
      Top             =   4515
      Width           =   2085
   End
   Begin VB.CheckBox CheckTest 
      Caption         =   "TEST"
      Height          =   345
      Left            =   5265
      TabIndex        =   49
      Top             =   4170
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdEOD 
      Caption         =   "EOD Processing"
      Height          =   585
      Left            =   255
      TabIndex        =   48
      Top             =   6975
      Width           =   2970
   End
   Begin VB.TextBox txtTerminalID 
      Height          =   540
      Left            =   5370
      TabIndex        =   45
      Top             =   6765
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CheckBox CheckDebit 
      Caption         =   "Allow Debit"
      Height          =   345
      Left            =   2940
      TabIndex        =   44
      Top             =   4170
      Width           =   2085
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   450
      Left            =   270
      TabIndex        =   43
      Top             =   8340
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtEventID 
      Height          =   510
      Left            =   9570
      TabIndex        =   41
      Top             =   6375
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtEventCode 
      Height          =   510
      Left            =   9525
      TabIndex        =   39
      Top             =   5790
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REPORTS"
      Height          =   840
      Left            =   270
      TabIndex        =   26
      Top             =   6060
      Width           =   3000
   End
   Begin VB.TextBox txtSuperPassword 
      Height          =   510
      Left            =   9450
      TabIndex        =   24
      Top             =   5175
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox txtUserName 
      Height          =   510
      Left            =   9420
      TabIndex        =   22
      Top             =   4530
      Visible         =   0   'False
      Width           =   1965
   End
   Begin Threed.SSCheck checkOnLine 
      Height          =   330
      Left            =   10635
      TabIndex        =   21
      Top             =   4095
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
      _Version        =   196609
      Caption         =   "Off-Line Transactions"
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   15
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPassword 
      Height          =   540
      Left            =   5370
      TabIndex        =   18
      Top             =   6090
      Width           =   1965
   End
   Begin VB.TextBox txtSellerCode 
      Height          =   540
      Left            =   5385
      TabIndex        =   17
      Top             =   5460
      Width           =   1965
   End
   Begin VB.TextBox txtIP 
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
      Left            =   5370
      TabIndex        =   15
      Text            =   "29.11.45.21"
      Top             =   4920
      Width           =   1530
   End
   Begin Threed.SSCheck CheckEvent 
      Height          =   390
      Left            =   495
      TabIndex        =   14
      Top             =   4830
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   688
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Use Test Event"
   End
   Begin VB.CheckBox CheckDebug 
      Caption         =   "DEBUG ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   510
      TabIndex        =   13
      Top             =   4170
      Width           =   2130
   End
   Begin Threed.SSCommand cmdTestInterface 
      Height          =   660
      Left            =   3405
      TabIndex        =   12
      Top             =   7635
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1164
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
      Caption         =   "Test Interface"
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   750
      Left            =   9900
      TabIndex        =   1
      Top             =   7590
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1323
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "EXIT"
   End
   Begin Threed.SSCommand cmdAddEdit 
      Height          =   675
      Left            =   10815
      TabIndex        =   0
      Top             =   8415
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1191
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Add/Edit Show Info"
   End
   Begin Threed.SSCommand cmdSetUpNet 
      Height          =   660
      Left            =   11340
      TabIndex        =   2
      Top             =   8415
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Setup Network"
   End
   Begin Threed.SSPanel panelShowData 
      Height          =   3975
      Left            =   45
      TabIndex        =   3
      Top             =   15
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox Check1 
         Caption         =   "TEST"
         Height          =   345
         Left            =   6345
         TabIndex        =   53
         Top             =   4260
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox ComboEvents 
         Height          =   450
         Left            =   2685
         TabIndex        =   36
         Text            =   "---SELECT---"
         Top             =   1515
         Width           =   6840
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "SUBMIT"
         Height          =   525
         Left            =   9840
         TabIndex        =   33
         Top             =   135
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtCity 
         Height          =   450
         Left            =   1935
         TabIndex        =   32
         Top             =   180
         Width           =   2970
      End
      Begin VB.ComboBox ComboMonth 
         Height          =   450
         Left            =   7140
         TabIndex        =   30
         Text            =   "---SELECT---"
         Top             =   180
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.ComboBox ComboEventCode 
         Height          =   450
         Left            =   2655
         TabIndex        =   28
         Text            =   "---SELECT---"
         Top             =   870
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select EVENT: "
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
         Index           =   7
         Left            =   135
         TabIndex        =   37
         Top             =   1575
         Width           =   2370
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EVENT MONTH"
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
         Index           =   6
         Left            =   5205
         TabIndex        =   31
         Top             =   255
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EVENT CATEGORY"
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
         Index           =   5
         Left            =   105
         TabIndex        =   29
         Top             =   930
         Visible         =   0   'False
         Width           =   2370
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EVENT CITY"
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
         Left            =   255
         TabIndex        =   27
         Top             =   210
         Width           =   1530
      End
      Begin VB.Label lblStopShow 
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
         Left            =   6285
         TabIndex        =   9
         Top             =   3435
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblStartShow 
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
         Left            =   2730
         TabIndex        =   8
         Top             =   3420
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblShowName 
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
         Left            =   2730
         TabIndex        =   7
         Top             =   2205
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show Name"
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
         Index           =   0
         Left            =   1065
         TabIndex        =   6
         Top             =   2235
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Index           =   1
         Left            =   1065
         TabIndex        =   5
         Top             =   3465
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Index           =   2
         Left            =   4725
         TabIndex        =   4
         Top             =   3495
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin Threed.SSCommand cmdBegin 
      Height          =   675
      Left            =   150
      TabIndex        =   10
      Top             =   7620
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1191
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
      Caption         =   "Single"
   End
   Begin Threed.SSCommand cmdTicketStock 
      Height          =   855
      Left            =   300
      TabIndex        =   11
      Top             =   5115
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1508
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Set/Reset Ticket Stock"
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   660
      Left            =   6525
      TabIndex        =   38
      Top             =   7650
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1164
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
      Caption         =   "Clear Sales History"
   End
   Begin Threed.SSCommand cmdMult 
      Height          =   675
      Left            =   1785
      TabIndex        =   47
      Top             =   7635
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1191
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
      Caption         =   "Mult"
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   3450
      TabIndex        =   56
      Top             =   8340
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "Coupon"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Terminal ID"
      Height          =   420
      Index           =   3
      Left            =   3585
      TabIndex        =   46
      Top             =   6855
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblSuper 
      Alignment       =   1  'Right Justify
      Caption         =   "Event ID: "
      Height          =   390
      Index           =   3
      Left            =   7710
      TabIndex        =   42
      Top             =   6450
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblSuper 
      Alignment       =   1  'Right Justify
      Caption         =   "Event Code: "
      Height          =   390
      Index           =   2
      Left            =   7665
      TabIndex        =   40
      Top             =   5865
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   2775
      TabIndex        =   35
      Top             =   8715
      Visible         =   0   'False
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
      Left            =   5205
      TabIndex        =   34
      Top             =   8640
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label lblSuper 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name: "
      Height          =   390
      Index           =   1
      Left            =   7590
      TabIndex        =   25
      Top             =   5250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblSuper 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name: "
      Height          =   390
      Index           =   0
      Left            =   7560
      TabIndex        =   23
      Top             =   4605
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   420
      Index           =   2
      Left            =   3585
      TabIndex        =   20
      Top             =   6180
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Seller Code"
      Height          =   420
      Index           =   1
      Left            =   3525
      TabIndex        =   19
      Top             =   5535
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Server IP"
      Height          =   420
      Index           =   0
      Left            =   3960
      TabIndex        =   16
      Top             =   4935
      Width           =   1350
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCommand As String
Dim rs As Recordset
Dim xxxx As Integer
Function ConstructBarCodeCoupon(TotalTicketNumber As Integer, ByRef thisTicketType As Integer) As String
    Dim i                       As Integer
    Dim j                       As Integer
    Dim K                       As Integer
    Dim strBarcodeString        As String
'        JComputerID = "A"
'        j = Len(Trim(JournalID))
'        strBarcodeString = "1"
'        If UCase(Trim(JComputerID)) = "A" Then strBarcodeString = "1"
'        If UCase(Trim(JComputerID)) = "B" Then strBarcodeString = "2"
'        If UCase(Trim(JComputerID)) = "C" Then strBarcodeString = "3"
'        If UCase(Trim(JComputerID)) = "D" Then strBarcodeString = "4"
'        If UCase(Trim(JComputerID)) = "E" Then strBarcodeString = "5"
'        For i = 1 To 9 - j
'             strBarcodeString = strBarcodeString & "0"
'        Next i
'        strBarcodeString = strBarcodeString & Trim(JournalID)
'        strBarcodeString = Left(JournalID, 2) & "Z" & Right(JournalID, 3)
            strBarcodeString = "801234567890"
        
'        strBarcodeString = "88" & Right(SellerCode, 2) & Format(Right(CStr(SessionID), 6), "00000#")
'        For i = 1 To 8
'            strBarcodeString = strBarcodeString & Trim(CStr(TotalTicketNumber))
'        Next i
        If TotalTicketNumber >= 99 Then TotalTicketNumber = 0
           TicketNumStr = CStr(TotalTicketNumber)
        If (Len(TicketNumStr) < 2) Then TicketNumStr = "0" & TicketNumStr
        strBarcodeString = strBarcodeString & TicketNumStr
   
        ConstructBarCodeCoupon = strBarcodeString



End Function




Sub PrintTicketsCoupon()
            Dim rsHeaders As Recordset
            Dim thisType As Integer
            MSComm1.CommPort = 4
            MSComm1.Settings = "9600,N,8,1"
            ' Open the port.
            MSComm1.PortOpen = True

            ''Set rsHeaders = db.OpenRecordset("Select * from TICKET_HEADERS where ShowID=" & ShowID)
            'If bDebug Then MsgBox strPrint
            
            strPrint = Chr(2) & Chr(11) & Chr(2) & Chr(11) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) ''    {clear printer buffer}
            'If bDebug Then MsgBox strPrint
            
            MSComm1.Output = strPrint
            ''''''''''''''NO BAR CODE '''''''''''''''''''''''''
'''            'strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8) & Chr(5)  ''    //init barcode Vertical
'''            If xxxx = 1 Then
'''                strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8)
'''            ElseIf xxxx = 2 Then
'''                strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2)
'''            ElseIf xxxx = 3 Then
'''                strPrint = Chr(2) & Chr(8) & Chr(1)
'''            ElseIf xxxx = 4 Then
'''                strPrint = Chr(2) & Chr(8)
'''            ElseIf xxxx = 5 Then
'''                strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8) & Chr(5)
'''            ElseIf xxxx = 6 Then
'''                strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8) & Chr(5)
'''            Else
'''                strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8) & Chr(5)  ''    //init barcode Vertical
'''            End If
'''            MSComm1.Output = strPrint
'''            '
'''            'thisType = 0
'''            strPrint = Chr(2) & Chr(16) & "COUPON"
'''            MSComm1.Output = strPrint
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
             strPrint = Chr(2) & Chr(16) & Chr(8) & " *AUTO SHOW STORE DISCOUNT*"

            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(2) & "$2 OFF"
            
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(11) & "N CONCOURSE & GALLERIA"
            
            MSComm1.Output = strPrint
               strPrint = Chr(2) & Chr(17) & Chr(34) & "  $2 OFF"
            
            MSComm1.Output = strPrint

            strPrint = Chr(2) & Chr(18) & Chr(2) & "2008NY"
            'If bDebug Then MsgBox strPrint
            
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(11) & "MINIMUM $15 PURCHASE"  'UCase(ln(2))"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(34) & "  2008NY"
            MSComm1.Output = strPrint
            
            strPrint = Chr(2) & Chr(19) & Chr(2) & "COUPON"
            MSComm1.Output = strPrint
            
            strPrint = Chr(2) & Chr(19) & Chr(11) & "******* $2 OFF ****** " 'UCase(ln(3))"
            MSComm1.Output = strPrint
            
                strPrint = Chr(2) & Chr(19) & Chr(34) & "  COUPON"
            MSComm1.Output = strPrint
            
            strPrint = Chr(2) & Chr(20) & Chr(2) & "CPN#1"
            MSComm1.Output = strPrint
            'strPrint = Chr(2) & Chr(20) & Chr(11) & "****MARCH 3 - 7, 2004**"  ''UCase(ln(4))
            strPrint = Chr(2) & Chr(20) & Chr(11) & "MARCH 21 - 30, 2008"  ''UCase(ln(4))
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(34) & "  CPN#1"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(2) & "K01"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(11) & Now()
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(34) & "  K01"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(1) & "2008NY"
            MSComm1.Output = strPrint
            'strPrint = Chr(2) & Chr(22) & Chr(10) & Format(Now(), "m/d/yyyy")
            strPrint = Chr(2) & Chr(22) & Chr(10) & "NOT TO BE COMBINED WITH"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(34) & "  2008NY"
            MSComm1.Output = strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(11) & "    ANY OTHER OFFERS"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(34) & "   A1234"
            MSComm1.Output = strPrint

           MSComm1.Output = Chr(2) & "pn"
    MSComm1.PortOpen = False

End Sub


Sub setupShow()
Dim rsInfo As Recordset

Set rsInfo = db.OpenRecordset("Select * from SHOWS where ShowID=" & ShowID)

'Set rsInfo = db.OpenRecordset("Select * from SHOWS where #" & Now() & "# between ShowStartDate and ShowEndDate")
'ShowID = rsInfo("ShowID")
'ShowID = 7
'lblShowName.Caption = rsInfo("ShowName")
'lblShowName.Caption = rsInfo("ShowName")
'lblStartShow.Caption = rsInfo("ShowStartDate")
'lblStopShow.Caption = rsInfo("ShowEndDate")
'lblTitles(5).Caption = "Pricing for " & Format(Now(), "MMM DD, YYYY")
'Set rsInfo = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategory,SHOW_TICKET_PRICES.* from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and (ShowDateStart between #" & Format(Now(), "mm/dd/yyyy") & " 12:00:00 AM# and #" & Format(Now(), "mm/dd/yyyy") & " 11:59:59 PM#) order by SHOW_TICKET_PRICES.ListSequence")
'With GridPrices
'    .Clear
'    .ColWidth(0) = 1000
'    .ColWidth(1) = 1000
'    .ColWidth(2) = 2000
'    .ColWidth(3) = 3000
'    .ColWidth(4) = 1000
'    .FixedAlignment(0) = 4
'    .FixedAlignment(1) = 4
'    .FixedAlignment(2) = 4
'    .FixedAlignment(3) = 4
'    .FixedAlignment(4) = 4
'    .ColAlignment(0) = 1
'    .ColAlignment(1) = 1
'    .ColAlignment(2) = 1
'    .ColAlignment(3) = 1
'    .ColAlignment(4) = 1
'    .Row = 0
'    .Col = 0
'    .Text = "Start"
'    .Col = 1
'    .Text = "Stop"
'    .Col = 2
'    .Text = "Category"
'    .Col = 3
'    .Text = "SubText"
'    .Col = 4
'    .Text = "Price"
'    Do Until rsInfo.EOF
'        .Rows = .Rows + 1
'        .Row = .Rows - 1
'        .Col = 0
'        .Text = Format(rsInfo("ShowDateStart"), "hh:mm AMPM")
'        .Col = 1
'        .Text = Format(rsInfo("ShowDateStop"), "hh:mm AMPM")
'        .Col = 2
'        .Text = rsInfo("TicketCategory")
'        .Col = 3
'        If Not IsNull(rsInfo("subButtonText")) Then .Text = rsInfo("subButtonText")
'        .Col = 4
'        .Text = Format(rsInfo("TicketPrice"), "$#.00")
'         rsInfo.MoveNext
'    Loop
'End With
End Sub


Private Sub AN_Click()
If AN.Value Then
   bAN = True
Else
    bAN = False
End If
End Sub

Private Sub CheckDebug_Click()
If CheckDebug.Value = 1 Then
    bDebug = True
Else
    bDebug = False

End If
End Sub

Private Sub CheckEvent_Click(Value As Integer)
If CheckEvent.Value = 0 Then
    ShowID = 9
Else
    ShowID = 7
End If
setupShow
End Sub

Private Sub checkOnLine_Click(Value As Integer)
If checkOnLine.Value = -1 Then
    lblSuper(0).Visible = True
    lblSuper(1).Visible = True
    lblSuper(2).Visible = True
    lblSuper(3).Visible = True
    txtUserName.Visible = True
    txtSuperPassword.Visible = True
    txtEventID.Visible = True
    txtEventCode.Visible = True
Else
    lblSuper(0).Visible = False
    lblSuper(1).Visible = False
    lblSuper(2).Visible = False
    lblSuper(3).Visible = False
    txtUserName.Visible = False
    txtSuperPassword.Visible = False
    txtEventID.Visible = False
    txtEventCode.Visible = False
End If
End Sub

Private Sub cmdAddEdit_Click()
frmSetUp.Show
End Sub

Private Sub cmdBatch_Click()
IP = txtIP.Text
frmBatch.Show
End Sub

Private Sub cmdBegin_Click()
Dim strTestEvent As String
Dim rs As Recordset
If ComboEvents.ListIndex < 0 Then
    MsgBox "Please select EVENT!"
    Exit Sub
End If
If CheckDebit.Value = 1 Then
    NotAllowDebit = False
Else
    NotAllowDebit = True
End If
If chkZipCode.Value = 1 Then
    GetZip = True
Else
    GetZip = fales
End If
If CheckTest.Value = 1 Then bNoPorts = True
'If Len(Trim(txtTerminalID.Text)) = 0 Then
'  MsgBox "Please enter TERMINALID"
'  Exit Sub
'Else
'    TerminalID = Trim(txtTerminalID.Text)
'End If

IP = txtIP.Text
SellerCode = txtSellerCode.Text
Password = txtPassword.Text
'If CheckEvent.Value = -1 Then
'    bTest = True
'    strTestEvent = "Y"
'Else
    bTest = False
    strTestEvent = "N"
'End If
If checkOnLine.Value = -1 Then
    Set rs = db.OpenRecordset("Select * from ADMINISTRATOR where username='" & LCase(txtUserName.Text) & "' and password='" & LCase(txtSuperPassword.Text) & "'")
    If rs.EOF Then
        MsgBox "Not a valid User Name and/or Password"
        Exit Sub
    Else
        If Len(Trim(txtEventCode.Text)) = 0 Or Len(Trim(txtEventID.Text)) = 0 Then
            MsgBox "Please enter BOTH and Event Code and Event ID"
            Exit Sub
        Else
            bOffLine = True
            EventID = txtEventID.Text & Space(10 - Len(txtEventID.Text))
            EventCode = txtEventCode.Text
        End If
    End If
Else
    bOffLine = False
End If
If CheckVISAMC.Value = 1 Then
    VisaMCOnly = True
Else
    VisaMCOnly = False
End If

If CheckDiscover.Value = 1 Then
    AllowDiscover = True
Else
    AllowDiscover = False
End If
    
If checkDEMO.Value = 1 Then
    bDEMO = True
End If

db.Execute "UPDATE INIT set IPAddress='" & IP & "',SellerCode='" & SellerCode & "',Password='" & Password & "',TestEvent='" & strTestEvent & "'"

Load frmPINPAD
frmWelcome.Show
Unload Me

End Sub

Private Sub cmdClear_Click()
Response = MsgBox("Are you sure that you want to clear SALES HISTORY?", vbYesNo + vbCritical)
If Response = vbYes Then
    db.Execute "Delete * from PURCHASE_INFO"
    db.Execute "Delete * from TICKET_CATEGORY_RECEIPTS"
    db.Execute "Delete * from LOG"
    db.Execute "Delete * from SESSION"
    

End If
End Sub

Private Sub cmdEOD_Click()
IP = txtIP.Text
SellerCode = txtSellerCode.Text
Password = txtPassword.Text

frmEOD.Show
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload frmPINPAD

End
End Sub

Private Sub cmdMult_Click()
If CheckDebit.Value = 1 Then
    NotAllowDebit = False
Else
    NotAllowDebit = True
End If
'If Len(Trim(txtTerminalID.Text)) = 0 Then
'  MsgBox "Please enter TERMINALID"
'  Exit Sub
'Else
'    TerminalID = Trim(txtTerminalID.Text)
'End If
If Len(Trim(txtCity.Text)) = 0 Then
    MsgBox "Please enter an EVENT CITY"
    Exit Sub
Else
    EventCity = txtCity.Text
End If
IP = txtIP.Text
SellerCode = txtSellerCode.Text
Password = txtPassword.Text
'If CheckEvent.Value = -1 Then
'    bTest = True
'    strTestEvent = "Y"
'Else
    bTest = False
    strTestEvent = "N"
'End If
If checkDEMO.Value = 1 Then
    bDEMO = True
End If

If checkOnLine.Value = -1 Then
    Set rs = db.OpenRecordset("Select * from ADMINISTRATOR where username='" & LCase(txtUserName.Text) & "' and password='" & LCase(txtSuperPassword.Text) & "'")
    If rs.EOF Then
        MsgBox "Not a valid User Name and/or Password"
        Exit Sub
    Else
        If Len(Trim(txtEventCode.Text)) = 0 Or Len(Trim(txtEventID.Text)) = 0 Then
            MsgBox "Please enter BOTH and Event Code and Event ID"
            Exit Sub
        Else
            bOffLine = True
            EventID = txtEventID.Text & Space(10 - Len(txtEventID.Text))
            EventCode = txtEventCode.Text
        End If
    End If
Else
    bOffLine = False
End If
If CheckVISAMC.Value = -1 Then
    VisaMCOnly = True
Else
    VisaMCOnly = False
End If

db.Execute "UPDATE INIT set IPAddress='" & IP & "',SellerCode='" & SellerCode & "',Password='" & Password & "',TestEvent='" & strTestEvent & "'"

bMultiple = True
Load frmPINPAD

frmWelcomeMult.Show
Unload Me
End Sub

Private Sub cmdNew_Click()
frmMain.Show

End Sub

Private Sub cmdSetUpNet_Click()
frmNetworkPath.Show
End Sub

Private Sub cmdSubmit_Click()
    SellerCode = Trim(txtSellerCode.Text)
    Password = Trim(txtPassword.Text)
    IP = Trim(txtIP.Text)
    Winsock1.RemoteHost = IP
    Winsock1.RemotePort = 1130
    strCommand = "Connect"
    Winsock1.Connect

End Sub

Private Sub cmdTestInterface_Click()
IP = txtIP.Text
SellerCode = txtSellerCode.Text
Password = txtPassword.Text
If CheckEvent.Value = ssCBChecked Then
    bTest = True
Else
    bTest = False
End If
frmTestInterface.Show

End Sub

Private Sub cmdTicketStock_Click()
frmTicketStock.Show 1
End Sub

Private Sub ComboCities_Change()

End Sub

Private Sub cmdZip_Click()
frmZipCode.Show 1
End Sub

Private Sub ComboEventCode_Click()
EventType = ComboEventCode.Text
strCommand = "SEVT"
Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:" & Trim(txtCity.Text) & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"

End Sub


Private Sub ComboEvents_Click()
rs.MoveFirst
rs.Move ComboEvents.ListIndex
ShowID = rs("ShowID")
EventType = rs("EventType")
EventCity = rs("EventCity")
EventReference = rs("EventReferenceName")
setupShow
'Set up Tickets
SellerCode = Trim(txtSellerCode.Text)
Password = Trim(txtPassword.Text)
IP = Trim(txtIP.Text)
'Response = MsgBox("Access TM Event data?", vbYesNo + vbInformation)
'If Response = vbYes Then
'    Winsock1.RemoteHost = IP
'    Winsock1.RemotePort = 1130
'    strCommand = "Connect"
'    Winsock1.Connect
'End If

End Sub


Private Sub Command1_Click()
''frmReport1.Show
frmMainAN.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
'Get computer data
'Dim ShowID As Integer
'Dim rs As Recordset
OpenDataBase
bTest = False
bDebug = False
bServer = True
bMultiple = False
bNoPorts = False
bDEMO = False
bAN = False
'ShowID = 9
Set rs = db.OpenRecordset("Select * from INIT")
'For j = 1 To 12
'    ComboMonth.AddItem CStr(j)
'Next j
txtSellerCode.Text = rs("SellerCode")
txtPassword.Text = rs("Password")
txtIP.Text = rs("IPAddress")

'txtSellerCode.Text = "ATMT01"
'txtPassword.Text = "ATMT01"
'txtIP.Text = "29.11.45.22"
Set rs = db.OpenRecordset("Select * from SHOWS where ShowEndDate >=#" & Now() & "#  order by ShowStartDate")
Do Until rs.EOF
    ComboEvents.AddItem rs("EventReferenceName")
    rs.MoveNext
Loop

'If rs("TestEvent") = "Y" Then
'  CheckEvent.Value = 1
'Else
'  CheckEvent.Value = 0
'End If
'setupShow
'Set rsInfo = db.OpenRecordset("Select * from SHOWS where #" & Now() & "# between ShowStartDate and ShowEndDate")
'ShowID = rsInfo("ShowID")
'ShowID = 7
'lblShowName.Caption = rsInfo("ShowName")
'lblShowName.Caption = rsInfo("ShowName")
'lblStartShow.Caption = rsInfo("ShowStartDate")
'lblStopShow.Caption = rsInfo("ShowEndDate")
'lblTitles(5).Caption = "Pricing for " & Format(Now(), "MMM DD, YYYY")
'Set rsInfo = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategory,SHOW_TICKET_PRICES.* from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and (ShowDateStart between #" & Format(Now(), "mm/dd/yyyy") & " 12:00:00 AM# and #" & Format(Now(), "mm/dd/yyyy") & " 11:59:59 PM#) order by SHOW_TICKET_PRICES.ListSequence")
'With GridPrices
'    .ColWidth(0) = 1000
'    .ColWidth(1) = 1000
'    .ColWidth(2) = 2000
'    .ColWidth(3) = 3000
'    .ColWidth(4) = 1000
'    .FixedAlignment(0) = 4
'    .FixedAlignment(1) = 4
'    .FixedAlignment(2) = 4
'    .FixedAlignment(3) = 4
'    .FixedAlignment(4) = 4
'    .ColAlignment(0) = 1
'    .ColAlignment(1) = 1
'    .ColAlignment(2) = 1
'    .ColAlignment(3) = 1
'    .ColAlignment(4) = 1
'    .Row = 0
'    .Col = 0
'    .Text = "Start"
'    .Col = 1
'    .Text = "Stop"
'    .Col = 2
'    .Text = "Category"
'    .Col = 3
'    .Text = "SubText"
'    .Col = 4
'    .Text = "Price"
'    Do Until rsInfo.EOF
'        .Rows = .Rows + 1
'        .Row = .Rows - 1
'        .Col = 0
'        .Text = Format(rsInfo("ShowDateStart"), "hh:mm AMPM")
'        .Col = 1
'        .Text = Format(rsInfo("ShowDateStop"), "hh:mm AMPM")
'        .Col = 2
'        .Text = rsInfo("TicketCategory")
'        .Col = 3
'        If Not IsNull(rsInfo("subButtonText")) Then .Text = rsInfo("subButtonText")
'        .Col = 4
'        .Text = Format(rsInfo("TicketPrice"), "$#.00")
'         rsInfo.MoveNext
'    Loop
'End With
End Sub


Private Sub SSCommand1_Click()
xxxx = CInt(txt_xxxx.Text)
PrintTicketsCoupon

End Sub


Private Sub Winsock1_Connect()

    strCommand = "LOGI"
    Winsock1.SendData "LOGI " & SellerCode & "    " & Password & "                                  " & Trim(EventCity) & "                      @#end#@"

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim strData2 As String
Dim header(5) As String
Select Case strCommand
Case "LOGI"
    strCommand = "SEVT"
    Winsock1.GetData strData
    If bDebug Then MsgBox "LOGI " & strData
    Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:" & EventCity & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"

    'logged in - get Event Types for this city
'    strCommand = "CATE"
'    Winsock1.GetData strData
'    If bDebug Then MsgBox "LOGI " & strData
'    Winsock1.SendData "CATE " & SellerCode & "    " & Password & "    " & Chr(10) & "City:" & Trim(EventCity) & Chr(10) & "@#end#@"
Case "CATE"
    'populate event type combo and wait for selection
    strCommand = ""
    Winsock1.GetData strData
    If bDebug Then MsgBox "CATE " & strData
    i = 1
    Do Until InStr(i, strData, Chr(10)) = 0
        ComboEventCode.AddItem Mid(strData, i, InStr(i, strData, Chr(10)) - i)
        i = InStr(i, strData, Chr(10)) + 1
    Loop
Case "SEVT"
       Winsock1.GetData strData
       If bDebug Then MsgBox "SEVT " & strData
       strCommand = "EVNT"
       'i = 1
        'Do Until InStr(i, strData, EventType) = 0
       '     i = InStr(i, strData, Chr(10))
        'Loop
       'i = InStr(1, strData, Chr(10))
       'EventID = Mid(strData, (i - 1) + 62, 10)
       'i = InStr(1, strData, "TORONTO TEST EVENT")
       i = InStr(1, strData, EventReference)
       EventID = Mid(strData, i + 40, 10)

       If bDebug Then MsgBox EventID
       Winsock1.SendData "EVNT " & SellerCode & "    " & Password & "    " & Chr(10) & "City: " & Trim(EventCity) & Chr(10) & "Event:" & EventID & "A" & Chr(10) & "@#end#@"
Case "EVNT"
     Winsock1.GetData strData
    If bDebug Then MsgBox "EVNT " & strData
   'Get Event Code
    EventCode = Mid(strData, InStr(1, strData, "EVENTCODE:") + 10, (InStr(1, strData, "LOCATION") - 1) - (InStr(1, strData, "EVENTCODE:") + 10))
    If Len(EventCode) < 19 Then
        EventCode = Trim(EventCode) & Space(19 - Len(Trim(EventCode)))
    Else
        EventCode = Left(EventCode, 19)
    End If
    If bDebug Then MsgBox "EventCode:" & EventCode
    'Get Show Name
    For i = 1 To 5
        header(i) = ""
    Next i
    i = 1
    j = 0
    Do Until InStr(i, strData, "HEADER:") = 0 Or j = 5
        i2 = InStr(InStr(i, strData, "HEADER:") + 1, strData, "HEADER:")
        If i2 > 0 Then
            j = j + 1
            header(j) = Mid(strData, InStr(i, strData, "HEADER:") + 7, i2 - (InStr(i, strData, "HEADER:") + 7))
        End If
        If bDebug Then MsgBox "Header(" & CStr(j) & ")=" & header(j)
        i = i2 + 1
    Loop
    EPDATE = Mid(strData, InStr(1, strData, "EPDATE") + 7, (InStr(1, strData, "DESC:") - 1) - (InStr(1, strData, "EPDATE") + 7))
    If bDebug Then MsgBox "EPDATE: " & EPDATE
    Location = Mid(strData, InStr(1, strData, "LOCATION:") + 10, (InStr(1, strData, "@#end") - 1) - (InStr(1, strData, "LOCATION:") + 10))
    If bDebug Then MsgBox "LOCATION: " & Location
   
    db.Execute "delete * from TICKET_HEADERS where ShowID=" & ShowID
    db.Execute "Insert into TICKET_HEADERS (ShowID, Header1,Header2,Header3,Header4,Header5,EPDATE,Location) values(" & ShowID & ",'" & header(1) & "','" & header(2) & "','" & header(3) & "','" & header(4) & "','" & header(5) & "','" & EPDATE & "','" & Location & "')"

End Select

End Sub


