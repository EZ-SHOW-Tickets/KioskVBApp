VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{7C6896B9-38B9-11D2-91BB-00C0A8356885}#1.0#0"; "DSINTX.OCX"
Begin VB.Form frmTestPinPad 
   Caption         =   "PIN Pad Test"
   ClientHeight    =   2340
   ClientLeft      =   3780
   ClientTop       =   3930
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   7305
   Begin VB.TextBox txtTerminalID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2610
      TabIndex        =   3
      Top             =   75
      Width           =   2190
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   90
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Debit Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   930
      TabIndex        =   0
      Top             =   630
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Terminal ID: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   885
      TabIndex        =   2
      Top             =   75
      Width           =   1710
   End
   Begin DSINTXLib.Dsintx Dsintx1 
      Left            =   135
      Top             =   915
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   609
      _StockProps     =   0
   End
   Begin VB.Label lblOutput 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3285
      TabIndex        =   1
      Top             =   1140
      Width           =   3570
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTestPinPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim CardReturn As String
If Len(Trim(txtTerminalID.Text)) = 0 Then
    MsgBox "Please enter TerminalID"
Else
    AccountNumber = ";4506300111103=16071200000000100100001?"
    Dsintx1.SetTfield "PinPadPort", 1
    Dsintx1.SetTfield "encrypt", "N"
    Dsintx1.SetTfield "PinPadPort", 1
    Dsintx1.SetTfield "IsAllegro", "Y"
    Dsintx1.SetTfield "MCredit", "Y"
    Dsintx1.SetTfield "IsDecoupled", "Y"
    Dsintx1.InitTransaction
    Dsintx1.SetTfield "Trans_Code", "01"
    Dsintx1.SetTfield "Track2_Acc", AccountNumber
    Dsintx1.SetTfield "Term_ID", txtTerminalID.Text
    Dsintx1.SetTfield "Amount_1", "50"
    CardReturn = Dsintx1.PrepareTransaction()
    MsgBox "SwitchCommString= " & CardReturn
    lblOutput.Caption = "ResponseCode=" & Dsintx1.GetTfield("Response_Code")
End If
End Sub


