VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmVerify 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9210
   ClientLeft      =   1815
   ClientTop       =   2370
   ClientWidth     =   13890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmVerify.frx":0000
   ScaleHeight     =   9210
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer ConvenienceFeeTimer 
      Left            =   465
      Top             =   1065
   End
   Begin VB.Timer Timer_Buy 
      Left            =   435
      Top             =   375
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   795
      Left            =   1200
      TabIndex        =   9
      Top             =   6960
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   1402
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
   Begin Threed.SSCommand cmdChange 
      Height          =   795
      Left            =   1200
      TabIndex        =   8
      Top             =   6060
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   1402
      _Version        =   196609
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Change"
   End
   Begin Threed.SSCommand cmdBuy 
      Height          =   1455
      Left            =   5850
      TabIndex        =   6
      Top             =   5715
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   2566
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   65535
      BackColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "BUY"
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   4050
      TabIndex        =   14
      Top             =   4425
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblConvenience2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press BUY to accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   5220
      Visible         =   0   'False
      Width           =   7590
   End
   Begin VB.Label lblConvenience1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE NOTE THAT A $1.50 CONVENIENCE FEE HAS BEEN ADDED."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3990
      TabIndex        =   12
      Top             =   5025
      Visible         =   0   'False
      Width           =   7590
   End
   Begin VB.Label XpressMember 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVerify.frx":2D72
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   4725
      TabIndex        =   11
      Top             =   7470
      Width           =   6150
   End
   Begin VB.Label lblFamilyPack 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVerify.frx":2E16
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   255
      TabIndex        =   10
      Top             =   3045
      Width           =   3525
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please verify the purchase information below and touch BUY to complete purchase."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1260
      Left            =   2100
      TabIndex        =   7
      Top             =   510
      Width           =   10140
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   4050
      TabIndex        =   5
      Top             =   3930
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4050
      TabIndex        =   4
      Top             =   3435
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   4050
      TabIndex        =   3
      Top             =   2940
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   4050
      TabIndex        =   2
      Top             =   2445
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   1950
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PURCHASE SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   4050
      TabIndex        =   0
      Top             =   1455
      Width           =   7260
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuy_Click()
Timer_Buy.Enabled = False

Call LogClick("VerifyTickets", "Buy")
If bConvenienceFee Then
    db.Execute "Insert into ConvenienceFee (SessionID,YorN) values(" & SessionID & ",'Y')"
End If
'frmPIN.Show
If bAN Then
    frmANAuthorization.Show
Else
    frmBankAuthorization.Show
End If
Unload Me
End Sub

Private Sub cmdCancel_Click()
Call LogClick("SpecifyTickets", "CancelTransaction")
If bConvenienceFee Then
    db.Execute "Insert into ConvenienceFee (SessionID,YorN) values(" & SessionID & ",'C')"
End If
'' CLEAR ALL TICKETS
Erase TicketTypes
Erase TotalTicketSold
TotTicketCount = 0
TicketCount = 0
'''''''''''''''''''

Timer_Buy.Enabled = False
'If bMultiple Then
'    frmWelcomeMult.Show
'Else
    'frmWelcome.Show
'    If UseAOP Then
        frmFirstPageAOP.Show
'    Else
'        frmFirstPage.Show
'    End If
'End If
Unload Me
End Sub

Private Sub cmdChange_Click()
Timer_Buy.interval = 0
Timer_Buy.Enabled = False
bChange = True
Call LogClick("VerifyTickets", "Change")
If bDebug Then MsgBox "Back to Specify"
frmSpecifyTickets.Show
Unload Me

End Sub

Private Sub ConvenienceFeeTimer_Timer()
ConvenienceFeeTimer.interval = 0
frmConvenienceFee.Show 1
If bAcceptConvenienceFee = False Then
   cmdCancel_Click
Else
    Timer_Buy.interval = 20000
    Timer_Buy.Enabled = True
End If
End Sub

Private Sub Form_Load()
'Dim totalprice As Currency
Dim gridRow As Integer
Dim TotalConvenienceFee As String
Timer_Buy.interval = 15000
Timer_Buy.Enabled = True
gridRow = 0
TotalPrice = 0
lblFamilyPack.Caption = ""
TotTicketCount = 0
If EventCity = "NEW YORK" Then
    XpressMember.Visible = True
Else
    XpressMember.Visible = False
End If

For i = 0 To UBound(TicketTypes, 2)
    If TotalTicketSold(i) > 0 Then
        ''For FAMILY PASS CHECK IF MORE THAN ONE
        If TicketTypes(1, i) = 15 And TotalTicketSold(i) > 1 Then
            lblFamilyPack.Caption = "Each single Family Pack admits 4 people. You Purchased " & CStr(TotalTicketSold(i)) & " Family Packs which admits " & CStr(4 * TotalTicketSold(i)) & " people.  If this isn't what you want please press Change Purchase above."
        End If
        gridRow = gridRow + 1
        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
        TotTicketCount = TotTicketCount + TotalTicketSold(i)
        lblGrid(gridRow).Caption = TotalTicketSold(i) & " " & TicketTypes(2, i) & " tickets @ " & Format(TicketTypes(3, i), "$#.00ea") & " = " & Format(TotalTicketSold(i) * TicketTypes(3, i), "$#.00")
        lblGrid(gridRow).Visible = True
    End If
Next i
If bConvenienceFee Then
        gridRow = gridRow + 1
        If ConvenienceFeeType = 1 Then
            TotalPrice = TotalPrice + ConvenienceFee
            lblGrid(gridRow).Caption = "Convenience Fee of " & FormatCurrency(ConvenienceFee, 2)
            TotalConvenienceFee = ConvenienceFee
        ElseIf ConvenienceFeeType = 2 Then
            TotalPrice = TotalPrice + ConvenienceFee * TotTicketCount
            lblGrid(gridRow).Caption = "Convenience Fee of " & FormatCurrency(ConvenienceFee * TotTicketCount, 2)
            TotalConvenienceFee = ConvenienceFee * TotTicketCount
        End If
        lblGrid(gridRow).ForeColor = &HFF&
        lblConvenience1.Caption = "PLEASE NOTE THAT A " & FormatCurrency(TotalConvenienceFee, 2) & " CONVENIENCE FEE HAS BEEN ADDED."
        lblConvenience1.Visible = True
        lblConvenience2.Visible = True
        
'        TotalPrice = TotalPrice + ConvenienceFee
'        lblGrid(gridRow).Caption = "Convenience Fee of $1.50"
        lblGrid(gridRow).Visible = True
End If
lblGrid(gridRow + 1).FontSize = 14
lblGrid(gridRow + 1).Caption = "Ticket Purchase Total: " & Format(TotalPrice, "$#.00")
lblGrid(gridRow + 1).Visible = True




'If bConvenienceFee Then
'    ConvenienceFeeTimer.Interval = 2
'Else
'    Timer_Buy.Interval = 20000
'    Timer_Buy.Enabled = True
'End If


End Sub






Private Sub Timer_Buy_Timer()
Timer_Buy.Enabled = False
cmdCancel_Click

End Sub


