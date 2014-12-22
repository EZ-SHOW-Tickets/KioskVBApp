VERSION 5.00
Begin VB.Form frmWelcomeMult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10005
   ClientLeft      =   -90
   ClientTop       =   1020
   ClientWidth     =   14925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmWelcomeMult.frx":0000
   ScaleHeight     =   0
   ScaleWidth      =   0
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   195
      Top             =   3750
   End
   Begin VB.Timer Timer2 
      Left            =   195
      Top             =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   195
      Top             =   2250
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   8
      Left            =   2190
      TabIndex        =   11
      Top             =   9210
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   7
      Left            =   2160
      TabIndex        =   10
      Top             =   8265
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   6
      Left            =   2160
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   5
      Left            =   2145
      TabIndex        =   8
      Top             =   6375
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   4
      Left            =   2160
      TabIndex        =   7
      Top             =   5430
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   4500
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   3555
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE SELECT THE EVENT FROM THE LIST BELOW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3255
      TabIndex        =   4
      Top             =   1125
      Width           =   8820
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   1
      Left            =   2145
      TabIndex        =   3
      Top             =   2610
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   630
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   1665
      Visible         =   0   'False
      Width           =   11760
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   12900
      Picture         =   "frmWelcomeMult.frx":44DC
      Top             =   240
      Width           =   1545
   End
   Begin VB.Image ImgStop 
      Height          =   1560
      Left            =   0
      Top             =   0
      Width           =   3090
   End
   Begin VB.Label LabelClose 
      BackStyle       =   0  'Transparent
      Height          =   810
      Left            =   300
      TabIndex        =   1
      Top             =   1770
      Width           =   765
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO BAYSIDE EXPO "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   525
      Index           =   0
      Left            =   3510
      TabIndex        =   0
      Top             =   345
      Width           =   8175
   End
End
Attribute VB_Name = "frmWelcomeMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMultShows As Recordset


Private Sub Form_Load()
Dim nShows As Integer
'Unload frmPINPAD
If EventCity = "BOSTON" Then
    lblWelcome(0).Caption = "WELCOME TO BAYSIDE EXPO"
ElseIf EventCity = "DETROIT" Then
    lblWelcome(0).Caption = "WELCOME TO COBO CENTER"
End If

nShows = 0
Set rsMultShows = db.OpenRecordset("Select * from SHOWS where EventCity='" & EventCity & "' and EventType <> 'TEST' and ShowEndDate > #" & Now & "# order by ShowStartDate")

Do Until rsMultShows.EOF
    lblShow(nShows).Caption = rsMultShows("ShowName") & " " & rsMultShows("ShowStartDate") & " - " & rsMultShows("ShowEndDate")
    lblShow(nShows).Visible = True
    rsMultShows.MoveNext
    nShows = nShows + 1
Loop
End Sub


Private Sub Image1_Click(Index As Integer)
'db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
'Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
'SessionID = rs(0)
'Call LogClick("Welcome", "SwipeCard")
'frmSpecifyTickets.Show
'Unload Me
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    AltDown = (Shift And vbAltMask) > 0
    If AltDown Then
        AccountNumber = ";4506300111103=16071200000000100100001?"
        CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
        CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
        CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
        CardExpire = Right(CardExpire, 2) & "/" & Left(CardExpire, 2)
        CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
        CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
        CardName = "Goodman/HS"
        bOffLine = True
        frmSpecifyTickets.Show
        Unload Me
    End If

End Sub


Private Sub ImgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    AltDown = (Shift And vbAltMask) > 0
    If AltDown Then
            frmMenu.Show
            Unload Me
    End If

End Sub


Private Sub LabelClose_DblClick()
End

End Sub



Private Sub lblShow_Click(Index As Integer)
Dim rs As Recordset
Dim IShow As Integer

rsMultShows.MoveFirst
rsMultShows.Move Index

Set rs = db.OpenRecordset("Select * from SHOWS where ShowID=" & rsMultShows("ShowID"))
ShowID = rs("ShowID")
EventType = rs("EventType")
EventCity = rs("EventCity")
EventReference = rs("EventReferenceName")
'setupShow
frmWelcome.Show
Unload Me
End Sub

Private Sub Timer2_Timer()
Dim xreturn As Boolean
Dim Track1 As String
On Error GoTo procError
xreturn = False
Timer2.Interval = 0
Timer2.Enabled = False
With frmPINPAD
    .Dsintx1.SetTfield "PinPadPort", 1
    .Dsintx1.SetTfield "NoWriteLog", "Y"
    .Dsintx1.InitTransaction
    .Dsintx1.SetTfield "ReadCard", "y"
    xreturn = .Dsintx1.ProcessTransaction
    If xreturn Then
        AccountNumber = .Dsintx1.GetTfield("Track2_Acc")
        If Len(AccountNumber) > 0 Then
            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            End If
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            Track1 = .Dsintx1.GetTfield("Track1")
            If Len(Track1) > 1 Then
                CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
            Else
                CardName = ""
            End If
            .Dsintx1.InitTransaction
            .Dsintx1.ReleaseConn
    '    'MsgBox CardNumber
    '    'MsgBox CardExpire
            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
            SessionID = rs(0)
            Call LogClick("Welcome", "SwipeCard")
            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
            'Me.Hide
            frmSpecifyTickets.Show
            Unload Me
        Else
            Timer2.Interval = 100
            Timer2.Enabled = True
        End If
    End If
    Exit Sub
End With
procError:
    Timer2.Interval = 100
    Timer2.Enabled = True
End Sub


Private Sub Timer3_Timer()
lblWelcome(3).ForeColor = &H8000&
lblWelcome(3).FontSize = 28
lblWelcome(3).Caption = "Swipe Credit Card to Begin"
Timer3.Interval = 0
Timer3.Enabled = False
Timer2.Interval = 100
Timer2.Enabled = True

End Sub


