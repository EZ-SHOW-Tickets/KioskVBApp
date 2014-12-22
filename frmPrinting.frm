VERSION 5.00
Begin VB.Form frmPrinting 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8760
   ClientLeft      =   675
   ClientTop       =   1950
   ClientWidth     =   11400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrinting.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerClick 
      Left            =   12075
      Top             =   150
   End
   Begin VB.Timer Timer_Exit 
      Left            =   12075
      Top             =   660
   End
   Begin VB.Label ClickLabel1 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   120
      TabIndex        =   5
      Top             =   7485
      Width           =   465
   End
   Begin VB.Label ClickLabel2 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   12030
      TabIndex        =   4
      Top             =   7320
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2805
      TabIndex        =   3
      Top             =   150
      Width           =   705
   End
   Begin VB.Label lblNumTickets 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   690
      TabIndex        =   2
      Top             =   2835
      Width           =   11820
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thank you for attending"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   4125
      Width           =   8985
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your tickets have printed in the tray below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   2925
      TabIndex        =   0
      Top             =   2175
      Width           =   8985
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClickClick1 As Boolean

Private Sub ClickLabel1_Click()
ClickClick1 = True
TimerClick.interval = 5000
TimerClick.Enabled = True
End Sub

Private Sub ClickLabel2_Click()
If ClickClick1 Then
        TimerClick.interval = 0
        TimerClick.Enabled = False
        ClickClick1 = False
        fTicketmation.Show 1
        
End If

End Sub

Private Sub Form_Load()
Dim nPrintCount As Integer
'Dim i As Integer
'Dim j As Integer
'Dim TotalPrice As Integer
'Dim rsTicketStock As Recordset
'Dim numTickets As Integer
'Dim LineNo(8) As Integer
'Dim FontSize(8) As Integer
'Dim RightSide As Integer
'
'numTickets = 0
'
'RightSide = 5100
'LineNo(1) = 200
'LineNo(2) = 500
'LineNo(3) = 900
'LineNo(4) = 1200
'LineNo(5) = 200
'LineNo(6) = 1600
'LineNo(7) = 1900
'LineNo(8) = 2300
'
'FontSize(1) = 14
'FontSize(2) = 12
'FontSize(3) = 12
'FontSize(4) = 14
'FontSize(5) = 12
'FontSize(6) = 10
'FontSize(7) = 12
'FontSize(8) = 12
'
'For i = 0 To UBound(TicketTypes, 2)
'    If TotalTicketSold(i) > 0 Then
'        With Printer
'            .Copies = TotalTicketSold(i)
'
'            .Orientation = 2
'
'            .Font = "Ariel Narrow"
'            .FontBold = True
'
'            .FontSize = FontSize(2)
'            .CurrentX = 50
'            .CurrentY = LineNo(2)
'            Printer.Print Format(TicketTypes(3, i), "$##.00") 'price
'
'            .FontSize = FontSize(3)
'            .CurrentX = 50
'            .CurrentY = LineNo(3)
'            Printer.Print "Price1"
'
'            .FontSize = FontSize(4)
'            .CurrentX = 50
'            .CurrentY = LineNo(4)
'            Printer.Print UCase(TicketTypes(2, i))
'
'            .FontSize = FontSize(6)
'            .CurrentX = 50
'            .CurrentY = LineNo(6)
'            Printer.Print strComputerName
'
'            .FontSize = FontSize(7)
'            .CurrentX = 50
'            .CurrentY = LineNo(7)
'            Printer.Print "04CIAS"
'
'            .FontSize = FontSize(1)
'            .CurrentX = 1100
'            .CurrentY = LineNo(1)
'            Printer.Print "** CANADIAN INTL AUTO SHOW **"
'
'            .FontSize = FontSize(2)
'            .CurrentX = 1500
'            .CurrentY = LineNo(2)
'            Printer.Print "FEB 13 - 22, 2004"
'
'            .FontSize = FontSize(3)
'            .CurrentX = 1200
'            .CurrentY = LineNo(3)
'            Printer.Print "VALID FOR ONE DAY ONLY"
'
'            .FontSize = 12
'            .CurrentX = 1200
'            .CurrentY = LineNo(4)
'            Printer.Print "CANADIAN INTL AUTO SHOW"
'            'Printer.Print "**2004 AUTO SHOW**"
'
'            .FontSize = FontSize(6)
'            .CurrentX = 1200
'            .CurrentY = LineNo(6)
'            Printer.Print Format(Now(), "MMM dd, yyyy") & " 10:30 AM"
'
'            .FontSize = FontSize(7)
'            .CurrentX = 1200
'            .CurrentY = LineNo(7)
'            Printer.Print Format(Now(), "mm/dd/yyyy") & " MC"
'
'            .FontSize = FontSize(8)
'            .CurrentX = 1200
'            .CurrentY = LineNo(8)
'            Printer.Print "VALID FOR FEB 13 - 22, 2004"
'
'            .FontSize = FontSize(2)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(2)
'            Printer.Print Format(TicketTypes(3, i), "$##.00")
'
'            .FontSize = FontSize(3)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(3)
'            Printer.Print "Price1"
'
'            .FontSize = FontSize(4)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(4)
'            Printer.Print UCase(TicketTypes(2, i))
'
'            .FontSize = FontSize(6)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(6)
'            Printer.Print strComputerName
'
'            .FontSize = FontSize(7)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(7)
'            Printer.Print "04CIAS"
'
'            .FontSize = FontSize(8)
'            .CurrentX = RightSide
'            .CurrentY = LineNo(8)
'            Printer.Print "AAAAAA"
'            .EndDoc
'        End With
'    End If
'Next i
'For i = 0 To UBound(TicketTypes, 2)
'    If TotalTicketSold(i) > 0 Then
'        TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
'    End If
'Next i
'
If bTicketmation Then
    bDEMO = False
    bTicketmation = False
End If

nPrintCount = 0
For i = 0 To UBound(TicketTypes, 2)
    If TicketTypes(1, i) = 8 Then
            'Family of 3
        nPrintCount = nPrintCount + 3 * TotalTicketSold(i)
    ElseIf TicketTypes(1, i) = 9 Then
            'Family of 4
        nPrintCount = nPrintCount + 4 * TotalTicketSold(i)
    ElseIf TicketTypes(1, i) = 14 Then
            'Family Pack
        nPrintCount = 4 * TotalTicketSold(i)
    ElseIf TicketTypes(1, i) = 13 Or TicketTypes(1, i) = 12 Or TicketTypes(1, i) = 11 Or TicketTypes(1, i) = 10 Then
            '2 Day Passes
        nPrintCount = nPrintCount + 2 * TotalTicketSold(i)
    ElseIf TicketTypes(1, i) = 18 Then
            '2 for 1
        nPrintCount = nPrintCount + 2 * TotalTicketSold(i)
'    ElseIf TicketTypes(1, i) = 16 Then
'            'PROGRAM (add coupon)
'        nPrintCount = nPrintCount + 2 * TotalTicketSold(i)
   
    Else
        nPrintCount = nPrintCount + TotalTicketSold(i)
End If

Next i

If PrinterError Then
    lblHeading(0).Caption = ""
    lblHeading(1).Caption = "PLEASE DO NOT USE THIS MACHINE"
    lblNumTickets.Caption = "THERE WAS A PRINTER PROBLEM.  PLEASE INFORM BOX OFFICE PERSONNEL."
    PrinterError = False
    bPrintError = True
    
Else
    lblNumTickets.Caption = "Please take " & CStr(nPrintCount) & " tickets from the tray below"
    If bDebug Then MsgBox "Timer set in Printing page"
    'Timer_Exit.interval = 4000
    Timer_Exit.interval = 3000
    Timer_Exit.Enabled = True
End If


End Sub





Private Sub Label1_Click()
    Unload Me
    frmMainAN.Show
End Sub


Private Sub Label2_Click()

End Sub

Private Sub Timer_Exit_Timer()
Dim rsTicketstock As Recordset

Timer_Exit.Enabled = False
If bDebug Then MsgBox "Timer1 Disabled"

'If bMultiple Then
'    frmWelcomeMult.Show
'Else
    On Error Resume Next
'    If UseAOP Then
        frmFirstPageAOP.Show
'    Else
'        frmFirstPage.Show
'    End If
    'frmWelcome.Show
    'Set rsTicketStock = db.OpenRecordset("Select * from TICKET_STOCK")
    'frmFirstPage.TXTbought.Caption = rsTicketStock(0)
    'frmFirstPage.Refresh
    'frmWelcome.TXTbought.Caption = rsTicketStock(0)
    'frmWelcome.Refresh
    'frmWelcome.Timer2.Interval = 1000
    'frmWelcome.Timer2.Enabled = True

'End If
Unload Me
End Sub


Private Sub TimerClick_Timer()
ClickClick1 = False
TimerClick.interval = 0

End Sub


