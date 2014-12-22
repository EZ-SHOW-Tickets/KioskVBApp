VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00B7B391&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9300
   ClientLeft      =   975
   ClientTop       =   2040
   ClientWidth     =   14160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14160
   WindowState     =   2  'Maximized
   Begin VB.Timer ClickTimer 
      Left            =   405
      Top             =   1950
   End
   Begin VB.Label clickImage2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1650
      Left            =   11265
      TabIndex        =   3
      Top             =   7680
      Width           =   2925
   End
   Begin VB.Label clickimage1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1200
      Left            =   0
      TabIndex        =   2
      Top             =   8130
      Width           =   2355
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
      Caption         =   "This machine is out of service at this time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3465
      Left            =   870
      TabIndex        =   0
      Top             =   2505
      Width           =   12255
   End
   Begin VB.Image ImgStop 
      Height          =   1335
      Left            =   240
      Picture         =   "frmStartup.frx":08CA
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imNo As Integer
Dim numcycles As Integer

Dim ClickClick1 As Boolean
Dim ClickClick2 As Boolean


Private Sub ClickImage1_Click()
ClickClick1 = True
ClickTimer.interval = 5000
ClickTimer.Enabled = True
End Sub


Private Sub ClickImage2_Click()
If ClickClick1 Then
        ClickTimer.interval = 0
        ClickTimer.Enabled = False
        ClickClick1 = False
        fTicketmation.Show 1
        'Unload Me
End If

End Sub

Private Sub ClickTimer_Timer()
ClickClick1 = False
ClickTimer.interval = 0

End Sub


Private Sub Form_Load()
Dim rsTicketstock As Recordset
Dim sappname As String


sappname = "Ticketmation"
'sAppPath = "D:\program files\Adobe\Photoshop 6.0\Photoshp.exe"
If App.PrevInstance Then
'If IsTaskRunning(sappname) Then
    MsgBox "Application '" & sappname & "' is running!"
    End
End If


If Not DBOpen Then
    OpenDataBase
    DBOpen = True
    ImgStop.Visible = False
End If
bStartup = True
Set rsTicketstock = db.OpenRecordset("Select * from TICKET_STOCK")
If rsTicketstock(0) <= 10 Then
    lblWelcome.Caption = "This Machine is OUT OF TICKET STOCK at this time"
Else
    lblWelcome.Caption = "This Machine is OUT OF SERVICE at this time"
End If
If AOPBarcode Then
    'ImgStop.Picture = LoadPicture(App.path & "/AOP_logo.gif")
    ImgStop = LoadPicture(App.path & "/AOP_logo.gif")
Else
    ImgStop.Picture = LoadPicture(App.path & "/ticketmation_logo_trans4.gif")
End If
End Sub

Private Sub ImgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    AltDown = (Shift And vbAltMask) > 0
    If AltDown Then
            'frmMenu.Show
            frmMainAN.Show
            Unload Me
    End If

End Sub


Private Sub Label1_Click()

End Sub

Private Sub LabelClose_DblClick()


End

End Sub



