VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1035
      Left            =   915
      TabIndex        =   0
      Top             =   585
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CAPTION_ATTACH = "Attach Device..."
Const CAPTION_CMD = "USB MSR Demo - Command Mode"
Const CAPTION_HID = "USB MSR Demo - Hid Mode"
Const CAPTION_IDLE = "USB MSR Demo - Searching"
Const CAPTION_KYBD = "USB MSR Demo - Keyboard Mode"
Const CAPTION_PROC = "Processing data..."
Const CAPTION_SWIPE = "Swipe card..."
Const STATE_DETERMINE_DEVICE = 1
Const STATE_HID_READ = 4
Const STATE_HID_START = 2
Const STATE_IDLE = 0
Const STATE_KEYBD_READ = 5
Const STATE_KEYBD_START = 3
Dim iCommandCount As Integer
Dim iState As Integer
Dim sCommandList() As String
Private MsgByte As Byte
Private MsgIndex As Long
Dim bBusy As Boolean
Dim bSuccess As Boolean
'
'************************
'* GoTotxtCommandEditor *
'************************------------------------------------------------------+
' Sets up display for input of next message.                                   |
'------------------------------------------------------------------------------+
Private Sub GoTotxtCommandEditor()
'  With txtCommandEditor
'    If .Visible Then
'      .SetFocus
'      .SelStart = 0
'      .SelLength = Len(HG_Command)
'    End If
'  End With
End Sub ' GoTotxtCommandEditor

'**********************
'* ProcessCommandFile *
'**********************--------------------------------------------------------+
'                                                                              |
'------------------------------------------------------------------------------+
Private Sub ProcessCommandFile()
  Dim fp As Integer
  Dim sCmd As String
  Dim sTmp As String

  cmdStartDownload.Enabled = False

  txtCommands.Text = ""
  ReDim sCommandList(255)
  iCommandCount = 0

  fp = FreeFile

  On Error GoTo GotError
  Open Trim(txtCmdFile.Text) For Input As #fp
  Do Until EOF(fp)
    Line Input #fp, sCmd

    sCmd = Trim(sCmd)

    txtCommands.Text = txtCommands.Text & sCmd & vbCrLf

    'delete comments
    If InStr(sCmd, ";") > 1 Then
      sCmd = Left$(sCmd, InStr(sCmd, ";") - 1)
    End If
    If InStr(sCmd, ";") = 1 Then
      sCmd = ""
    End If

    sCmd = Trim(sCmd)

    ' if there is data then get that command
    If Len(sCmd) > 1 Then
      sTmp = sFormatCommand(sCmd)
      If Len(sTmp) > 0 Then
        sCommandList(iCommandCount) = sTmp
        iCommandCount = iCommandCount + 1
      End If
    End If
  Loop

  Close #fp

  If iCommandCount = 0 Then
    cmdStartDownload.Enabled = False
    lblVerifyResults.Caption = "No commands found in file"
  Else
    cmdStartDownload.Enabled = True
    lblVerifyResults.Caption = "Found " & iCommandCount & " commands"
  End If

  Exit Sub
GotError:
  bSuccess = False
  lblVerifyResults.Caption = "Unable to read in " & txtCmdFile.Text
End Sub ' ProcessCommandFile


'
'******************
'* sFormatCommand *
'******************------------------------------------------------------------+
'                                                                              |
'------------------------------------------------------------------------------+
Private Function sFormatCommand(sMsg As String) As String
  Dim i As Integer
  Dim sTmp As String

  ' first strip all spaces
  sMsg = Replace(sMsg, " ", "")

  ' if command is "Delay" (case-insensitive) then exit with it
  If InStr(1, sMsg, "Delay", vbTextCompare) = 1 Then
    sFormatCommand = "Delay " & Mid$(sMsg, 6)
    Exit Function
  End If

  ' if any non-hex values found, then exit with null ("")
  For i = 1 To Len(sMsg)
    If InStr(1, "0123456789ABCDEF", Mid$(sMsg, i, 1), vbTextCompare) = 0 Then
      sFormatCommand = ""
      Exit Function
    End If
  Next i

  ' then add in a spaces for each byte ("xxyyzz" to "xx yy zz")
  For i = 1 To Len(sMsg) Step 2
    sTmp = sTmp & Mid$(sMsg, i, 2) & " "
  Next i

  ' remove the trailing space
  sTmp = Trim(sTmp)

  'and return the value
  sFormatCommand = sTmp

End Function ' sFormatCommand
'
'*******************
'* ClrDlgBtn_Click *
'*******************-----------------------------------------------------------+
'                                                                              |
'------------------------------------------------------------------------------+
Private Sub ClrDlgBtn_Click()
    txtCommands.Text = ""
End Sub ' ClrDlgBtn_Click


'****************
'* InterpretMsg *
'****************--------------------------------------------------------------+
' Checks the syntax of a message and converts the message into a format that   |
' can be sent to the device                                                    |
'------------------------------------------------------------------------------+
Public Function InterpretMsg(sText As String) As Boolean
  Dim NextByteResult As String
  Dim DataIndex As Long
  Dim FtrRptIndex As Long

  InterpretMsg = True
  MsgIndex = 1

  For FtrRptIndex = 1 To Capabilities.FeatureReportByteLength
    FtrRptTrxBfr(FtrRptIndex - 1) = 0
  Next

  If GetNextMsgByte(sText) = "GotByte" Then
    FtrRptTrxBfr(FRI_CMD) = MsgByte
  Else
    HandleBadMsg
    InterpretMsg = False
    Exit Function
  End If

  ' is user is adding their own length then skip over it now
  If chkAutoAddLen.Value = 0 Then
    MsgIndex = MsgIndex + 3
  End If

  DataIndex = FRI_DATA
  NextByteResult = "GotByte"
  Do While NextByteResult = "GotByte"
    NextByteResult = GetNextMsgByte(sText)
    If NextByteResult = "GotByte" Then
      If DataIndex = Capabilities.FeatureReportByteLength Then
        MsgBox ("Data portion of message can't be longer than " & Capabilities.FeatureReportByteLength - 3 & " bytes")
        InterpretMsg = False
        Exit Function
      End If
      FtrRptTrxBfr(DataIndex) = MsgByte
      DataIndex = DataIndex + 1
    ElseIf NextByteResult = "BadMsg" Then
      HandleBadMsg
      InterpretMsg = False
      Exit Function
    End If
  Loop

  If chkAutoAddLen.Value = 1 Then
    FtrRptTrxBfr(FRI_LEN) = DataIndex - FRI_DATA
  Else
    FtrRptTrxBfr(FRI_LEN) = Val("&h" & Mid(txtCommandEditor.Text, 4, 2)) 'DataIndex - FRI_DATA
  End If
End Function ' InterpretMsg
'
'****************
'* HandleBadMsg *
'****************--------------------------------------------------------------+
' Displays a message indicating a bad message syntax has been entered.         |
'------------------------------------------------------------------------------+
Public Sub HandleBadMsg()
  MsgBox ("Bad Message.  Message should be hexidecimal in format ""CMND [DATA]""." & vbCrLf & "For Example, ""00 00"".")
End Sub ' HandleBadMsg

'
'******************
'* GetNextMsgByte *
'******************------------------------------------------------------------+
' Gets the next message byte from a string at MsgIndex position and stores it  |
' in MsgByte                                                                   |
'------------------------------------------------------------------------------+
Private Function GetNextMsgByte(sText As String) As String
  Dim GotChar As Boolean
  Dim TmpByte As Byte
  Dim AscVal As Byte

  GetNextMsgByte = "MsgDone"
  MsgByte = 0
  GotChar = False

  Do While MsgIndex <= Len(sText)
    AscVal = Asc(Mid(sText, MsgIndex, 1))
    MsgIndex = MsgIndex + 1
    If AscVal = 9 Or AscVal = 10 Or AscVal = 13 Or AscVal = 32 Then
      If GotChar Then
        GetNextMsgByte = "GotByte"
        Exit Function
      End If
    ElseIf (AscVal >= 48 And AscVal <= 57) Or (AscVal >= 65 And AscVal <= 72) Or (AscVal >= 97 And AscVal <= 102) Then
      If AscVal >= 65 And AscVal <= 72 Then
        TmpByte = AscVal - 55
      ElseIf AscVal >= 97 And AscVal <= 102 Then
        TmpByte = AscVal - 87
      Else
        TmpByte = AscVal - 48
      End If

      If GotChar Then
        MsgByte = MsgByte * 16 + TmpByte
        GetNextMsgByte = "GotByte"
        Exit Function
      Else
        GotChar = True
        MsgByte = TmpByte
      End If
    Else
      GetNextMsgByte = "BadMsg"
      Exit Function
    End If
  Loop

  If GotChar Then
    GetNextMsgByte = "GotByte"
  End If
End Function ' GetNextMsgByte

'
'*************
'* ByteToHex *
'*************-----------------------------------------------------------------+
' Converts a byte to a hexidecimal string for display purposes.                |
'------------------------------------------------------------------------------+
Public Function ByteToHex(ByVal ByteVal As Byte) As String
  ByteToHex = Hex(ByteVal)
  If Len(ByteToHex) = 1 Then
    ByteToHex = "0" & ByteToHex
  End If
End Function ' ByteToHex

'
'*********************
'* ByteArrayToHexStr *
'*********************---------------------------------------------------------+
' Translates a byte array to a hexidecimal string for display purposes.        |
'------------------------------------------------------------------------------+
Private Function ByteArrayToHexStr(ByRef ByteArray() As Byte, ByVal ArrayLen As Long) As String
  Dim i As Long
  Dim HexStr As String

  ByteArrayToHexStr = ""
  For i = 1 To ArrayLen
    HexStr = Hex(ByteArray(FRI_DATA + i - 1))

    If Len(HexStr) = 1 Then
      HexStr = "0" & HexStr
    End If

    If i > 1 Then
      ByteArrayToHexStr = ByteArrayToHexStr & " " & HexStr
    Else
      ByteArrayToHexStr = HexStr
    End If
  Next i
End Function ' ByteArrayToHexStr
Private Sub Command1_Click()
  HG_Command = sFormatCommand("02")

  'Check for any fatal errors.
  If FatalError <> "" Then
    MsgBox ("Fatal Error - " & FatalError)
    Exit Sub
  End If

  'See if the correct device has been detected
  If HidDevice = INVALID_HANDLE_VALUE Then

    'See if the correct device is attached
    DetectDevice
    If HidDevice = INVALID_HANDLE_VALUE Then
      MsgBox ("Attach device before sending commands")
      Exit Sub
    End If
    If FatalError <> "" Then
      MsgBox ("Fatal Error - " & FatalError)
      Exit Sub
    End If
  End If

  'Initialize command data
  If Not InitCmd Then
    GoTotxtCommandEditor
    Exit Sub
  End If

  ' Interpret message set up the transmit data
  If Not InterpretMsg(HG_Command) Then
    GoTotxtCommandEditor
    Exit Sub
  End If

  ' Display Request
  'DisplayComDlg txtCommands, "Request"

  ' Send / Receive message
  SendCmd

  ' Display Response
  'DisplayComDlg txtCommands, "Response"

  GoTotxtCommandEditor
End Sub


