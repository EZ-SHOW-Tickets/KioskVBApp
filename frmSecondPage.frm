VERSION 5.00
Object = "{158336E7-3FF3-456E-912C-5985E9BBED24}#1.2#0"; "MTUSBHIDSwipe.ocx"
Begin VB.Form frmSecondPage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ctlUSBHID.USBHID USBHID2 
      Left            =   12090
      Top             =   330
      _ExtentX        =   1508
      _ExtentY        =   2037
   End
   Begin VB.Image Image1 
      Height          =   11475
      Left            =   525
      Picture         =   "frmSecondPage.frx":0000
      Stretch         =   -1  'True
      Top             =   195
      Width           =   14790
   End
End
Attribute VB_Name = "frmSecondPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
USBHID2.PortOpen = True
If Not (USBHID2.PortOpen) Then
 MsgBox "Couldn't open HID Swipe Reader"
 End
End If

Image1.Top = 0
'Image1.Height = Me.Height
'Image1.Width = Me.Width
End Sub

Private Sub USBHID2_CardDataChanged()
Dim abTrack1() As Byte
Dim abTrack2() As Byte
Dim abTrack3() As Byte
Dim i As Integer
Dim strOutput As String
''Text1.Text = ""
With USBHID2
If (.GetTrack(1) <> "" Or .GetTrack(2) <> "" Or .GetTrack(3) <> "") Then
        SwipTime = Now()
''
        AccountNumber = .GetTrack(2)
        If Len(AccountNumber) > 0 Then
            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
''
            If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                'Card is a debit
                'lblWelcome(3).FontSize = 16
                'lblWelcome(3).ForeColor = &HFF&
                'lblBadSwipe.Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                lblBadSwipe.Visible = True
                'Timer3.Interval = 3000
                'Timer3.Enabled = True
                Exit Sub
            'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
            ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
                'Card is an AMEX
                'lblWelcome(3).FontSize = 16
                'lblWelcome(3).ForeColor = &HFF&
                If NotAllowDebit Then
                    'lblBadSwipe.Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
                    lblBadSwipe.Visible = True
                Else
                    'lblBadSwipe.Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
                    lblBadSwipe.Visible = True
                End If
                'Timer3.Interval = 3000
                'Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
                'Card is a debit
                'lblWelcome(3).FontSize = 16
                'lblWelcome(3).ForeColor = &HFF&
                lblBadSwipe.Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
                'Timer3.Interval = 3000
                'Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
                'lblWelcome(3).FontSize = 16
                'lblWelcome(3).ForeColor = &HFF&
                lblBadSwipe.Caption = "PLEASE USE A VALID CARD."
                'Timer3.Interval = 3000
                'Timer3.Enabled = True
                Exit Sub
            End If
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            Track1 = .GetTrack(1)
            If bDebug Then MsgBox "Track1=" & Track1
            If bDebug Then MsgBox "Track2=" & AccountNumber
            If Len(Track1) > 1 Then
                CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
            Else
                CardName = .GetLName & "/" & .GetFName
            End If

            USBHID2.PortOpen = False

            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
            SessionID = rs(0)
            Call LogClick("Welcome", "SwipeCard")
            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
            frmSpecifyTickets.Show
            Unload Me

        End If

End If
End With


End Sub


Private Sub USBHID2_CardDataError()
MsgBox "BAD"
End Sub


