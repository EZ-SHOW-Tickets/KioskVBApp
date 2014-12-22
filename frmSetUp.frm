VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSetUp 
   BorderStyle     =   0  'None
   ClientHeight    =   9930
   ClientLeft      =   -465
   ClientTop       =   2085
   ClientWidth     =   15285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtComPort 
      Height          =   405
      Left            =   9135
      TabIndex        =   38
      Top             =   195
      Width           =   555
   End
   Begin Threed.SSFrame framePriceData 
      Height          =   3555
      Left            =   360
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   6271
      _Version        =   196609
      Begin VB.TextBox txtSubText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         TabIndex        =   36
         Top             =   1230
         Width           =   4440
      End
      Begin VB.ComboBox ComboShowDate 
         Height          =   315
         Left            =   2025
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   360
         Width           =   2910
      End
      Begin MSFlexGridLib.MSFlexGrid gridPriceTimes 
         Height          =   1155
         Left            =   225
         TabIndex        =   32
         Top             =   1890
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Rows            =   3
         Cols            =   24
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         FillStyle       =   1
      End
      Begin VB.ComboBox ComboTicketCategory 
         Height          =   315
         Left            =   2025
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   840
         Width           =   2910
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   450
         TabIndex        =   37
         Top             =   1245
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   390
         TabIndex        =   35
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Width           =   1530
      End
   End
   Begin Threed.SSPanel panelShowData 
      Height          =   5505
      Left            =   495
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9710
      _Version        =   196609
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   6120
         Pattern         =   "*.jpg"
         TabIndex        =   43
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtShowCity 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1875
         TabIndex        =   41
         Top             =   960
         Width           =   2550
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   7935
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3000
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   7935
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   2595
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   7935
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   2175
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   7935
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   1770
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2955
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   3360
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2955
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2950
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2955
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   2540
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2955
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2130
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame FrameTicketTypes 
         Caption         =   "Select all ticket types and basic prices for this show (can be edited later)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   8580
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   9
            Left            =   7485
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   1935
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   2490
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   300
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   330
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   105
            TabIndex        =   17
            Top             =   735
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   105
            TabIndex        =   16
            Top             =   1095
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   105
            TabIndex        =   15
            Top             =   1485
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   105
            TabIndex        =   14
            Top             =   1935
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   4860
            TabIndex        =   13
            Top             =   375
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   4860
            TabIndex        =   12
            Top             =   750
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   4860
            TabIndex        =   11
            Top             =   1170
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   4860
            TabIndex        =   10
            Top             =   1590
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.CheckBox CheckTicketTypes 
            Caption         =   "TicketType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   4860
            TabIndex        =   9
            Top             =   1995
            Visible         =   0   'False
            Width           =   2370
         End
      End
      Begin VB.TextBox txtEndDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5505
         TabIndex        =   5
         Top             =   1515
         Width           =   1710
      End
      Begin VB.TextBox txtBeginDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1875
         TabIndex        =   4
         Top             =   1500
         Width           =   1710
      End
      Begin VB.TextBox txtShowName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1875
         TabIndex        =   2
         Top             =   450
         Width           =   7140
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   480
         Left            =   360
         TabIndex        =   19
         Top             =   4800
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   847
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
         Caption         =   "SAVE"
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   480
         Left            =   2760
         TabIndex        =   39
         Top             =   4800
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   847
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
         Caption         =   "CANCEL"
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show Logo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4440
         TabIndex        =   42
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3915
         TabIndex        =   7
         Top             =   1545
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   255
         TabIndex        =   6
         Top             =   1500
         Width           =   1530
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   255
         TabIndex        =   3
         Top             =   465
         Width           =   1530
      End
   End
   Begin VB.ComboBox ComboShows 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   225
      Width           =   7920
   End
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTickets As Recordset
Dim ShowID As Integer
Dim StartStop As Integer

Private Sub CheckTicketTypes_Click(Index As Integer)
If CheckTicketTypes(Index).Value = 1 Then
    txtPrice(Index).Visible = True
Else
    txtPrice(Index).Visible = False
End If
End Sub

Private Sub cmdCancel_Click()
'frmMenu.Show
frmMainAN.Show
Unload Me

End Sub

Private Sub cmdSave_Click()
Dim j As Integer
Dim rsNewShow As Recordset
db.Execute "Insert into SHOWS (ShowName,ShowStartDate,ShowEnddate) values('" & txtShowName.text & "',#" & txtBeginDate.text & "#,#" & txtEndDate.text & "#)"
Set rsNewShow = db.OpenRecordset("Select max(showID) as newShowID from SHOWS")
ShowID = rsNewShow("newShowID")
For i = 1 To DateDiff("d", txtBeginDate.text, txtEndDate.text)
    'Find ticket Types and prices
    ComboShowDate.AddItem Format(DateAdd("d", i - 1, txtBeginDate.text), "MMM DD,YYYY")
    For j = 0 To CheckTicketTypes.Count - 1
        If CheckTicketTypes(j).Value = 1 Then
            rsTicketTypes.MoveFirst
            rsTicketTypes.Move j
            db.Execute "Insert into SHOW_TICKET_PRICES (ShowID,ShowDateStart,TicketCategoryID,TicketPrice,ListSequence,ShowDateStop) values(" & ShowID & ",#" & DateAdd("d", i - 1, txtBeginDate.text) & "#," & rsTicketTypes("TicketCategoryID") & "," & Right(txtPrice(j).text, Len(txtPrice(j).text) - 1) & "," & j & ",#" & DateAdd("d", i - 1, txtBeginDate.text) & " 11:59:59 PM#)"
        End If
    Next j
Next i
Set rsShows = db.OpenRecordset("Select * from SHOWS where ShowEndDate > #" & Now() & "#")
With gridPriceTimes
For i = 0 To 23
    .Row = 0
    .ColWidth(i) = 500
    .Col = i
    If i > 11 Then
        .text = Format(i - 11, "#PM")
    Else
        .text = Format(i + 1, "#AM")
    End If
Next i
    .Col = 8
    .Row = 1
    .text = "start"
    .Col = 21
    .Row = 1
    .text = "stop"
End With
framePriceData.Visible = True
End Sub

Private Sub ComboShowDate_Click()

'Set rsInfo = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategory,SHOW_TICKET_PRICES.* from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and (ShowDateStart between #" & Format(Now(), "mm/dd/yyyy") & " 12:00:00 AM# and #" & Format(Now(), "mm/dd/yyyy") & " 11:59:59 PM#) order by SHOW_TICKET_PRICES.ListSequence")
'With GridPrices
'    .ColWidth(0) = 1000
'    .ColWidth(1) = 1000
'    .ColWidth(2) = 2000
'    .ColWidth(3) = 3000
'    .ColWidth(4) = 1000
'    .FixedAlignment(0) = 4
'    .FixedAlignment(1) = 4
'    .FixedAlignment(2) = 4
'    .FixedAlignment(3) = 4
'    .FixedAlignment(4) = 4
'    .ColAlignment(0) = 1
'    .ColAlignment(1) = 1
'    .ColAlignment(2) = 1
'    .ColAlignment(3) = 1
'    .ColAlignment(4) = 1
'    .Row = 0
'    .Col = 0
'    .Text = "Start"
'    .Col = 1
'    .Text = "Stop"
'    .Col = 2
'    .Text = "Category"
'    .Col = 3
'    .Text = "SubText"
'    .Col = 4
'    .Text = "Price"
'    Do Until rsInfo.EOF
'        .Rows = .Rows + 1
'        .Row = .Rows - 1
'        .Col = 0
'        .Text = Format(rsInfo("ShowDateStart"), "hh:mm AMPM")
'        .Col = 1
'        .Text = Format(rsInfo("ShowDateStop"), "hh:mm AMPM")
'        .Col = 2
'        .Text = rsInfo("TicketCategory")
'        .Col = 3
'        If Not IsNull(rsInfo("subButtonText")) Then .Text = rsInfo("subButtonText")
'        .Col = 4
'        .Text = Format(rsInfo("TicketPrice"), "$#.00")
'         rsInfo.MoveNext
'    Loop
'End With

End Sub


Private Sub ComboShows_Click()
Dim i
If ComboShows.ListIndex > 0 Then
    rsShows.MoveFirst
    rsShows.Move ComboShows.ListIndex - 1
    ShowID = rsShows("ShowID")
    txtShowName.text = rsShows("Showname")
    txtBeginDate.text = rsShows("ShowStartDate")
    txtEndDate.text = rsShows("ShowEndDate")
    Set rsTickets = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategoryID,TicketCategory,TicketPrice,subButtonText from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " order by SHOW_TICKET_PRICES.ListSequence")
    Do Until rsTickets.EOF
        rsTicketTypes.MoveFirst
        i = -1
        Do Until rsTicketTypes.EOF
            i = i + 1
            If rsTicketTypes("TicketCategoryID") = rsTickets("TicketCategoryID") And CheckTicketTypes(i).Value <> 1 Then
                CheckTicketTypes(i).Value = 1
                ComboTicketCategory.AddItem rsTickets("TicketCategory")
                txtPrice(i).text = Format(rsTickets("TicketPrice"), "$#.00")
            End If
            rsTicketTypes.MoveNext
        Loop
        rsTickets.MoveNext
    Loop
    For i = 1 To DateDiff("d", txtBeginDate.text, txtEndDate.text)
        ComboShowDate.AddItem Format(DateAdd("d", i - 1, txtBeginDate.text), "MMM DD,YYYY")
    Next i
    ComboShowDate.ListIndex = 0
    ComboTicketCategory.ListIndex = 0
    framePriceData.Visible = True
Else

End If
    panelShowData.Visible = True

End Sub


Private Sub ComboTicketCategory_Click()
Dim rs As Recordset
rsTicketTypes.MoveFirst
rsTicketTypes.Move ComboTicketCategory.ListIndex
Set rs = db.OpenRecordset("Select * from SHOW_TICKET_PRICES Where ShowID=" & ShowID & " and TicketCategoryID=" & rsTicketTypes("TicketCategoryID") & " and #" & ComboShowDate.text & "# between ShowDateStart and ShowDateStop")
If Not IsNull(rs("subButtonText")) Then txtSubText.text = rs("subButtonText")
With gridPriceTimes
    For i = 0 To 23
        .Row = 0
        .ColWidth(i) = 500
        .Col = i
        If i > 11 Then
            .text = Format(i - 11, "#PM")
        Else
            .text = Format(i + 1, "#AM")
        End If
    Next i
    Do Until rs.EOF
        i = Hour(Format(rs("ShowDateStart"), "hh:mm"))
        .Col = i
        .Row = 1
        .text = "start"
        i = Hour(Format(rs("ShowDateStop"), "hh:mm"))
        .Col = i
        .Row = 1
        .text = "stop"
        rs.MoveNext
    Loop
End With
framePriceData.Visible = True
End Sub


Private Sub Form_Load()
Dim i As Integer
StartStop = 3
Set rsShows = db.OpenRecordset("Select * from SHOWS where ShowEndDate > #" & Now() & "#")
ComboShows.AddItem "Add new show"
Do Until rsShows.EOF
    ComboShows.AddItem rsShows("ShowName")
    rsShows.MoveNext
Loop
ComboShows.text = "---Select---"
Set rsTicketTypes = db.OpenRecordset("Select * from TICKET_CATEGORIES")
i = -1
Do Until rsTicketTypes.EOF
    i = i + 1
    CheckTicketTypes(i).Caption = rsTicketTypes("TicketCategory")
    CheckTicketTypes(i).Visible = True
    rsTicketTypes.MoveNext
Loop

End Sub


Private Sub gridPriceTimes_Click()

With gridPriceTimes
    If .Row = 1 Then
        If Len(.text) > 0 Then
            If .text = "start" And StartStop = 3 Then
                StartStop = 2
            ElseIf .text = "start" And StartStop = 1 Then
                StartStop = 0
            ElseIf .text = "stop" And StartStop = 3 Then
                StartStop = 1
            ElseIf .text = "stop" And StartStop = 2 Then
                StartStop = 0
            End If
            .text = ""
        End If
    ElseIf .Row = 2 Then
        .Row = 1
        If .text = "start" Or .text = "X" Then
            .Row = 2
            .text = ""
        End If
    End If

End With
End Sub


Private Sub gridPriceTimes_DblClick()
    With gridPriceTimes
        If .Row = 1 Then

                If StartStop = 3 Then
                    .text = "X"
                    .Row = 2
                    .text = "$0.00"
                ElseIf StartStop = 2 Then
                    .text = "start"
                     StartStop = 3
                    .Row = 2
                    .text = "$0.00"
                 ElseIf StartStop = 1 Then
                    .text = "stop"
                    StartStop = 3
                Else
                    .text = "start"
                    StartStop = 1
                     .Row = 2
                    .text = "$0.00"
                End If
        End If
    End With

End Sub


Private Sub gridPriceTimes_KeyPress(KeyAscii As Integer)
With gridPriceTimes
    If .Row = 2 Then
        .Row = 1
        If .text = "start" Or .text = "X" Then
            .Row = 2
            If KeyAscii = 13 Then
                 .text = Format(.text, "#.00")
            Else
                .text = .text & Chr(KeyAscii)
            End If
        End If
    End If
End With
End Sub

Private Sub txtComPort_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 intComPort = CInt(txtComPort.text)
End If
End Sub


Private Sub txtPrice_GotFocus(Index As Integer)
txtPrice(Index).text = Right(txtPrice(Index).text, Len(txtPrice(Index).text) - 1)
End Sub

Private Sub txtPrice_LostFocus(Index As Integer)
txtPrice(Index).text = Format(txtPrice(Index).text, "$#.00")
End Sub


