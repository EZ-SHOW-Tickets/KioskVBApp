VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmTicketStock 
   Caption         =   "Ticket Stock"
   ClientHeight    =   4215
   ClientLeft      =   4635
   ClientTop       =   3225
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7725
   Begin Threed.SSOption PinPadTarget 
      Height          =   330
      Index           =   0
      Left            =   4275
      TabIndex        =   17
      Top             =   1920
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Stock"
      Value           =   -1
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      TabIndex        =   15
      Top             =   1050
      Width           =   2760
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   3390
      TabIndex        =   14
      Top             =   2445
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   2640
      TabIndex        =   13
      Top             =   2445
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   1920
      TabIndex        =   12
      Top             =   2445
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   1200
      TabIndex        =   11
      Top             =   2445
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   2445
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1935
      TabIndex        =   7
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton cmdKeyPad 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5805
      TabIndex        =   4
      Top             =   1935
      Width           =   1650
   End
   Begin VB.TextBox txtTicketStock 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3210
      TabIndex        =   2
      Top             =   255
      Width           =   1815
   End
   Begin Threed.SSCommand cmdCalcel 
      Height          =   780
      Left            =   630
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1376
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
   End
   Begin Threed.SSCommand cmdReset 
      Height          =   780
      Left            =   3030
      TabIndex        =   1
      Top             =   3270
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1376
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reset"
   End
   Begin Threed.SSOption PinPadTarget 
      Height          =   330
      Index           =   1
      Left            =   4275
      TabIndex        =   18
      Top             =   2385
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tickets"
   End
   Begin VB.Label lblTicketStock 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting Ticket Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   16
      Top             =   1110
      Width           =   3015
   End
   Begin VB.Label lblTicketStock 
      Alignment       =   1  'Right Justify
      Caption         =   "Remaining Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Top             =   390
      Width           =   3015
   End
End
Attribute VB_Name = "frmTicketStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcel_Click()
Unload Me
End Sub


Private Sub cmdClear_Click()
txtTicketStock.text = ""
txtNumber.text = ""
End Sub


Private Sub cmdKeyPad_Click(Index As Integer)
Dim i As Integer
If Index = 9 Then
  i = 0
Else
    i = Index + 1
End If
If PinPadTarget(0).Value Then
    txtTicketStock.text = txtTicketStock.text & CStr(i)
Else
    txtNumber.text = txtNumber.text & CStr(i)
End If
End Sub


Private Sub cmdReset_Click()
'Update database
If Len(Trim(txtTicketStock.text)) > 0 And Len(Trim(txtNumber.text)) > 0 Then
    db.Execute "Update TICKET_STOCK set TicketStock=" & txtTicketStock.text & ",StartingTix=" & txtNumber.text
    Unload Me

Else
    MsgBox "Please enter data for both stock and ticket number"
End If
'frmWelcome.TXTbought.Caption = txtTicketStock.Text
'frmFirstPage.TXTbought.Caption = txtTicketStock.Text
End Sub


Private Sub Form_Load()
Dim rsTicketstock As Recordset
Set rsTicketstock = db.OpenRecordset("Select * from TICKET_STOCK")
txtTicketStock.text = rsTicketstock(0)
txtNumber.text = rsTicketstock(1)
PinPadTarget(0).Value = True

End Sub


Private Sub lblTicketStock_Click(Index As Integer)
    PinPadTarget(Index).Value = True
    
End Sub


