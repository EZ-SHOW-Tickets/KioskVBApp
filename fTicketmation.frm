VERSION 5.00
Begin VB.Form fTicketmation 
   Caption         =   "Ticketmation"
   ClientHeight    =   2205
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4215
   Begin VB.Timer TimerTimeOut 
      Left            =   180
      Top             =   1380
   End
   Begin VB.CommandButton bLogEnter 
      BackColor       =   &H00808080&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   960
      TabIndex        =   3
      Top             =   1170
      Width           =   2430
   End
   Begin VB.CommandButton bLog3 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2910
      TabIndex        =   2
      Top             =   180
      Width           =   1080
   End
   Begin VB.CommandButton bLog2 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1590
      TabIndex        =   1
      Top             =   180
      Width           =   1080
   End
   Begin VB.CommandButton Blog1 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "fTicketmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BPick(3) As Integer
Dim Picked As Integer

Private Sub Command1_Click()

End Sub


Private Sub Blog1_Click()
If Picked < 3 Then
    Picked = Picked + 1
    BPick(Picked) = 1
End If
End Sub

Private Sub bLog2_Click()
If Picked < 3 Then
    Picked = Picked + 1
    BPick(Picked) = 2
End If

End Sub


Private Sub bLog3_Click()
If Picked < 3 Then
    Picked = Picked + 1
    BPick(Picked) = 3
End If

End Sub


Private Sub bLogEnter_Click()

If BPick(1) = 1 And BPick(2) = 3 And BPick(3) = 1 And Not bStartup Then
    ' 1-3-1  From WELCOME - Run with no card swipe
    bDEMO = True
    bTicketmation = True
    'Set next session
    db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
    Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
    SessionID = rs(0)

    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    Unload Me
    frmSpecifyTickets.Show

    'If UseAOP Then
        Unload frmFirstPageAOP
    'Else
    '    Unload frmFirstPage
    'End If
    'frmSpecifyTickets.Show

ElseIf BPick(1) = 1 And BPick(2) = 1 And BPick(3) = 1 And Not bStartup And Not bPrintError Then
    ' 1-1-1 From WELCOME - OUT
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    Unload Me
    'frmMenu.Show
    frmMainAN.Show
    'If UseAOP Then
        Unload frmFirstPageAOP
    'Else
    '    Unload frmFirstPage
    'End If
    'frmMainAN.Show
ElseIf BPick(1) = 1 And BPick(2) = 1 And BPick(3) = 1 And bPrintError Then
    ' 1-1-1 From WELCOME - OUT
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    Unload Me
    Unload frmPrinting
    'frmMenu.Show
    frmMainAN.Show
    'If UseAOP Then
    'Else
    '    Unload frmFirstPage
    'End If
    'frmMainAN.Show
ElseIf BPick(1) = 1 And BPick(2) = 2 And BPick(3) = 1 And Not bStartup Then
    ' 1-2-1 From WELCOME - to Ticketstock
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    TimerTimeOut.Enabled = False
    frmTicketStock.Show 1
    'If UseAOP Then
        Unload frmFirstPageAOP
    'Else
    '    Unload frmFirstPage
    'End If
    Unload Me
    'If UseAOP Then
        frmFirstPageAOP.Show
    'Else
    '   frmFirstPage.Show
    'End If
'ElseIf BPick(1) = 3 And BPick(2) = 3 And BPick(3) = 3 And Not bStartup Then
'    ' 3-3-3 From WELCOME - EXIT
'    BPick(1) = 0
'    BPick(2) = 0
'    BPick(3) = 0
'    TimerTimeOut.Enabled = False
'    End
ElseIf BPick(1) = 1 And BPick(2) = 2 And BPick(3) = 3 And Not bStartup Then
    ' 1-2-3 From WELCOME - to View Log
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    TimerTimeOut.Enabled = False
    frmLOG.Show 1
'    If UseAOP Then
'        Unload frmFirstPageAOP
'    Else
'        Unload frmFirstPage
'    End If
    Unload Me

'ElseIf BPick(1) = 1 And BPick(2) = 1 And BPick(3) = 1 And Not bStartup Then
'    'From WELCOME - RESET
'    BPick(1) = 0
'    BPick(2) = 0
'    BPick(3) = 0
'    Unload frmWelcome
'    frmWelcome.Show
'    Unload Me
ElseIf BPick(1) = 1 And BPick(2) = 1 And BPick(3) = 1 And bStartup Then
    'From STARTUP - to Menu
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    bDEMO = True
    bStartup = False
    'Moved here for test
    Unload Me
    frmMainAN.Show
    Unload frmStartup
    'Unload Me
   ' frmMainAN.Show
ElseIf BPick(1) = 3 And BPick(2) = 3 And BPick(3) = 3 And Not bStartup Then
    'From STARTUP - to Menu
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    'Moved here for test
    Unload Me
    frmSecurityTix.Show 1
    Unload Me
   ' frmMainAN.Show

Else
    'No Code is correct
    BPick(1) = 0
    BPick(2) = 0
    BPick(3) = 0
    Picked = 0
    Unload Me
    'Go back to where we came from
'    If bStartup Then
'       frmStartup.Show
'    Else
'       frmWelcome.Show
'    End If
End If

TimerTimeOut.Enabled = False
End Sub


Private Sub Form_Load()
BPick(1) = 0
BPick(2) = 0
BPick(3) = 0

Picked = 0
TimerTimeOut.interval = 10000
TimerTimeOut.Enabled = True
End Sub


Private Sub Timer1_Timer()

End Sub


Private Sub TimerTimeOut_Timer()
Unload Me
End Sub


