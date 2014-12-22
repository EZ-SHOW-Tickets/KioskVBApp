VERSION 5.00
Object = "{158336E7-3FF3-456E-912C-5985E9BBED24}#1.2#0"; "MTUSBHIDSwipe.ocx"
Begin VB.Form frmCCSwipe 
   Caption         =   "CCSwipe"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5970
      Left            =   1830
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmSwipe.frx":0000
      Top             =   195
      Width           =   7665
   End
   Begin ctlUSBHID.USBHID USBHID1 
      Left            =   345
      Top             =   360
      _ExtentX        =   1667
      _ExtentY        =   1482
   End
End
Attribute VB_Name = "frmCCSwipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
USBHID1.PortOpen = True
If Not (USBHID1.PortOpen) Then
 MsgBox "Couldn't open HID Swipe Reader"
 End
End If

End Sub


Private Sub USBHID1_CardDataChanged()
Dim abTrack1() As Byte
Dim abTrack2() As Byte
Dim abTrack3() As Byte
Dim i As Integer
Dim strOutput As String
Text1.Text = ""
With USBHID1
If (.GetTrack(1) <> "" Or .GetTrack(2) <> "" Or .GetTrack(3) <> "") Then
    
    Text1.Text = Text1.Text & "CardData = " & .CardData
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "CardDataMasked = " & .CardDataMasked
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "FirstName = " & .GetFName
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "LastName = " & .GetLName
    Text1.Text = Text1.Text & vbCrLf
    
    Text1.Text = Text1.Text & "Track1Data = " & .GetTrack(1)
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track2Data = " & .GetTrack(2)
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track3Data = " & .GetTrack(3)
    Text1.Text = Text1.Text & vbCrLf
    
    Text1.Text = Text1.Text & "Track1DataLength = " & .Track1Length
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track2DataLength = " & .Track2Length
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track3DataLength = " & .Track3Length
    Text1.Text = Text1.Text & vbCrLf
    
    Text1.Text = Text1.Text & "FirstNameMasked = " & .GetFNameMasked
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "LastNameMasked = " & .GetLNameMasked
    Text1.Text = Text1.Text & vbCrLf
    
    Text1.Text = Text1.Text & "Track1DataLengthmasked = " & .Track1LengthMasked
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track2DataLengthmasked = " & .Track2LengthMasked
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track3DataLengthmasked = " & .Track3LengthMasked
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track1DataMasked = " & .GetTrackMasked(1)
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track2DataMasked = " & .GetTrackMasked(2)
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Track3DataMasked = " & .GetTrackMasked(3)
    Text1.Text = Text1.Text & vbCrLf
    
    Text1.Text = Text1.Text & "DUKPTKSN = " & .DUKPTKSN
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "SessionID = " & .SessionID
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "SessionIDHex = " & .SessionIDHexString
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "MagnePrint = " & .MPrintData
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "MagnePrintHex = " & .MagnePrintDataHexString
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "MagnePrint Length = " & .MPrintLen
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "Sequence Number = " & .SequenceNumber
    Text1.Text = Text1.Text & vbCrLf
    Text1.Text = Text1.Text & "ReaderID = " & .ReaderID
    Text1.Text = Text1.Text & vbCrLf
End If
End With

End Sub


Private Sub USBHID1_CardDataError()
MsgBox ("Error Reading Card")

End Sub


