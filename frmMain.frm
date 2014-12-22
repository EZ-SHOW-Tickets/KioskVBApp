VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "WALKUP SYSTEMS"
   ClientHeight    =   6060
   ClientLeft      =   1755
   ClientTop       =   3375
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComboEventCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2745
      TabIndex        =   11
      Text            =   "---SELECT---"
      Top             =   1965
      Width           =   2430
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   150
      Top             =   5385
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   2010
      TabIndex        =   7
      Text            =   "29.11.45.21"
      Top             =   240
      Width           =   1530
   End
   Begin VB.TextBox txtSellerCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5625
      TabIndex        =   6
      Text            =   "ATMT01"
      Top             =   120
      Width           =   1965
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9810
      TabIndex        =   5
      Text            =   "ATMT01"
      Top             =   120
      Width           =   1965
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10575
      TabIndex        =   4
      Top             =   1560
      Width           =   1710
   End
   Begin VB.ComboBox ComboMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9915
      TabIndex        =   2
      Text            =   "---SELECT---"
      Top             =   990
      Width           =   2430
   End
   Begin VB.TextBox txtEventCity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2025
      TabIndex        =   0
      Text            =   "BOSTON"
      Top             =   945
      Width           =   5205
   End
   Begin VB.Label lblTitles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EVENT MONTH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   300
      TabIndex        =   12
      Top             =   1980
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Server IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   255
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Seller Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   3765
      TabIndex        =   9
      Top             =   195
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   8025
      TabIndex        =   8
      Top             =   210
      Width           =   1755
   End
   Begin VB.Label lblTitles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EVENT MONTH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   7470
      TabIndex        =   3
      Top             =   1005
      Width           =   2400
   End
   Begin VB.Label lblTitles 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EVENT CITY:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   1005
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubmit_Click()
    SellerCode = Trim(txtSellerCode.Text)
    Password = Trim(txtPassword.Text)
    IP = Trim(txtIP.Text)
    Winsock1.RemoteHost = IP
    Winsock1.RemotePort = 1130
    strCommand = "Connect"
    Winsock1.Connect

End Sub

Private Sub Form_Load()
Dim MonthNames(12) As String
Dim i As Integer

OpenDataBase
bTest = False
bDebug = True
MonthNames(1) = "Jan"
MonthNames(2) = "Feb"
MonthNames(3) = "Mar"
MonthNames(4) = "Apr"
MonthNames(5) = "May"
MonthNames(6) = "Jun"
MonthNames(7) = "Jul"
MonthNames(8) = "Aug"
MonthNames(9) = "Sep"
MonthNames(10) = "Oct"
MonthNames(11) = "Nov"
MonthNames(12) = "Dec"
For i = 1 To 12
    ComboMonth.AddItem MonthNames(i)

Next i

End Sub


Private Sub Winsock1_Connect()
    If bDebug Then MsgBox "Connected"
    strCommand = "LOGI"
    SellerCode = Trim(txtSellerCode.Text)
    Password = Trim(txtPassword.Text)
    EventCity = Trim(txtEventCity.Text)
    'strData = "LOGI " & SellerCode & "    " & Password & "                                  BOSTON                       @#end#@"
   'strData = "LOGI " & SellerCode & "    " & Password & "                                 " & EventCity & Space(29 - Len(EventCity)) & "@#end#@"
    strData = "LOGI " & SellerCode & "    " & Password & "                                  BOSTON                       @#end#@"
    If bDebug Then MsgBox strData
    Winsock1.SendData strData
    'Winsock1.SendData "LOGI " & SellerCode & "    " & Password & "                                 " & Trim(EventCity) & Space(29 - Len(Trim(txtEventCity.Text))) & "@#end#@"


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim strData2 As String
Dim header(5) As String
Select Case strCommand
Case "LOGI"
'    strCommand = "SEVT"
'    Winsock1.GetData strData
'    If bDebug Then MsgBox "LOGI " & strData
'    Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:" & EventCity & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"

    'logged in - get Event Types for this city
    Winsock1.GetData strData
    If bDebug Then MsgBox "LOGI " & strData
    strCommand = "CATE"
    strData = "CATE " & SellerCode & "    " & Password & "    " & Chr(10) & "City:" & EventCity & Chr(10) & "@#end#@"
    If bDebug Then MsgBox strData
    Winsock1.SendData strData
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


