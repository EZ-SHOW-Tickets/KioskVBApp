VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmMainAN 
   Caption         =   "Ticketmation Admin"
   ClientHeight    =   9135
   ClientLeft      =   1875
   ClientTop       =   1395
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSecurity 
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
      Height          =   420
      Left            =   4440
      TabIndex        =   43
      Text            =   "6"
      Top             =   2280
      Width           =   645
   End
   Begin VB.TextBox txtSkip 
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
      Height          =   420
      Left            =   8130
      TabIndex        =   40
      Text            =   "0"
      Top             =   2640
      Width           =   1125
   End
   Begin VB.CheckBox CheckEMail 
      Caption         =   "Get EMail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   39
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSessionCounter 
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
      Left            =   8010
      TabIndex        =   37
      Text            =   "1"
      Top             =   2040
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test No Response"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox CheckBackup 
      Caption         =   "Use Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   34
      Top             =   1320
      Width           =   1695
   End
   Begin Threed.SSCommand cmdDeviceUP 
      Height          =   495
      Left            =   8760
      TabIndex        =   33
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "+"
   End
   Begin VB.CheckBox CheckAOPBarcode 
      Caption         =   "AOP BarCode"
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
      Left            =   240
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check_AOP 
      Caption         =   "AOP Screen"
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
      Left            =   240
      TabIndex        =   30
      Top             =   2280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton OptionPrinter 
      Caption         =   "PA"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   29
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton OptionPrinter 
      Caption         =   "BOCA"
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   28
      Top             =   4200
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Threed.SSCommand cmdPracAutoPrint 
      Height          =   735
      Left            =   9000
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Practical Automation"
   End
   Begin Threed.SSCommand cmdSetuoEvent 
      Height          =   855
      Left            =   7920
      TabIndex        =   26
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Download SuperApp"
   End
   Begin VB.CheckBox CheckReceipt 
      Caption         =   "Print Receipt Question"
      Height          =   345
      Left            =   7800
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.OptionButton OptionCFee 
      Caption         =   "Per Ticket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6000
      TabIndex        =   24
      Top             =   8355
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.OptionButton OptionCFee 
      Caption         =   "Per Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6000
      TabIndex        =   23
      Top             =   7995
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtConvenienceFee 
      Height          =   450
      Left            =   4080
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   7995
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox chkConvenienceFee 
      Caption         =   "Charge Convenience Fee"
      Height          =   345
      Left            =   195
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   3900
   End
   Begin Threed.SSCommand cmdLOG 
      Height          =   615
      Left            =   2400
      TabIndex        =   19
      Top             =   4680
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "View Log"
   End
   Begin VB.CheckBox chkZipCode 
      Caption         =   "Get ZIP Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   18
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Threed.SSCommand cmdBatch 
      Height          =   615
      Left            =   5640
      TabIndex        =   17
      Top             =   3360
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "Ticket Print"
   End
   Begin VB.CheckBox AN 
      Caption         =   "AN"
      Height          =   345
      Left            =   9435
      TabIndex        =   16
      Top             =   8040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox CheckDiscover 
      Caption         =   "Allow Discover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   14
      Top             =   1365
      Value           =   1  'Checked
      Width           =   2310
   End
   Begin VB.CheckBox CheckVISAMC 
      Caption         =   "No AMEX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   13
      Top             =   1785
      Width           =   2085
   End
   Begin VB.CheckBox checkDEMO 
      Caption         =   "DEMO ONLY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   12
      Top             =   1845
      Width           =   2085
   End
   Begin VB.CheckBox CheckDebit 
      Caption         =   "Allow Debit"
      Height          =   345
      Left            =   9420
      TabIndex        =   10
      Top             =   8415
      Visible         =   0   'False
      Width           =   2085
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
      Left            =   8040
      TabIndex        =   5
      Text            =   "1"
      Top             =   1320
      Width           =   645
   End
   Begin VB.CheckBox CheckDebug 
      Caption         =   "DEBUG ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   180
      TabIndex        =   4
      Top             =   1470
      Width           =   2130
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   975
      Left            =   255
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "EXIT"
   End
   Begin Threed.SSPanel panelShowData 
      Height          =   1230
      Left            =   75
      TabIndex        =   1
      Top             =   -45
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2170
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox Check1 
         Caption         =   "TEST"
         Height          =   345
         Left            =   6345
         TabIndex        =   15
         Top             =   4260
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox ComboEvents 
         Height          =   450
         Left            =   2685
         TabIndex        =   7
         Text            =   "---SELECT---"
         Top             =   315
         Width           =   6840
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select EVENT: "
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
         Left            =   135
         TabIndex        =   8
         Top             =   405
         Width           =   2370
      End
   End
   Begin Threed.SSCommand cmdBegin 
      Height          =   1125
      Left            =   360
      TabIndex        =   2
      Top             =   4680
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1984
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
      Caption         =   "START"
   End
   Begin Threed.SSCommand cmdTicketStock 
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1508
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Set/Reset Ticket Stock"
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   660
      Left            =   2400
      TabIndex        =   9
      Top             =   3960
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1164
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
      Caption         =   "Clear Sales History"
   End
   Begin Threed.SSCommand cmdMult 
      Height          =   675
      Left            =   7080
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1191
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
      Caption         =   "Mult"
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   2415
      TabIndex        =   21
      Top             =   5400
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "Batch"
   End
   Begin Threed.SSCommand cmdCoupon 
      Height          =   735
      Left            =   6360
      TabIndex        =   32
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Test Email"
   End
   Begin Threed.SSCommand cmdSessionCounter 
      Height          =   495
      Left            =   8730
      TabIndex        =   36
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "+"
   End
   Begin Threed.SSCommand cmdSkip 
      Height          =   495
      Left            =   9330
      TabIndex        =   41
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "+"
   End
   Begin Threed.SSCommand cmdSecurity 
      Height          =   495
      Left            =   5160
      TabIndex        =   44
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "+"
   End
   Begin VB.Label lblSecurity 
      Alignment       =   1  'Right Justify
      Caption         =   "Security Cutoff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   45
      Top             =   2400
      Width           =   1950
   End
   Begin VB.Label lblBarcode 
      Alignment       =   1  'Right Justify
      Caption         =   "Skip These Barcodes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5520
      TabIndex        =   42
      Top             =   2760
      Width           =   2550
   End
   Begin VB.Label lblSession 
      Alignment       =   1  'Right Justify
      Caption         =   "Session Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   38
      Top             =   2160
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Device ID"
      Height          =   420
      Index           =   0
      Left            =   6630
      TabIndex        =   6
      Top             =   1335
      Width           =   1350
   End
End
Attribute VB_Name = "frmMainAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCommand As String
Dim rs As Recordset

Sub SetAutoNumber(sTable As String, ByVal lNum As Long)
On Error GoTo Err_SetAutoNumber
    ' Purpose:      set the AutoNumber field in sTable to begin at lNum.
    ' Arguments:    sTable = name of table to modify.
    '               lNum = the number you wish to begin from.
    ' Sample use:   Call SetAutoNumber("tblInvoice", 1000)
    'Dim db As DAO.Database      ' Current db.
    Dim tdf As DAO.TableDef     ' TableDef of sTable.
    Dim i As Integer            ' Loop counter
    Dim fld As DAO.Field        ' Field of sTable.
    Dim sFieldName As String    ' Name of the AutoNumber field.
    Dim vMaxID As Variant       ' Current Maximum AutoNumber value.
    Dim sSQL As String          ' Append/Delete query string.
    Dim sMsg As String          ' MsgBox string.

    lNum = lNum - 1             ' Assign to 1 less than desired value.

    ' Locate the auto-incrementing field for this table.
    'Set db = CurrentDb()
    Set tdf = db.TableDefs(sTable)
    For i = 0 To tdf.Fields.Count - 1
        Set fld = tdf.Fields(i)
        If fld.Attributes And dbAutoIncrField Then
            sFieldName = fld.Name
            Exit For
        End If
    Next

    If Len(sFieldName) = 0 Then
        sMsg = "No AutoNumber field found in table """ & sTable & """."
        MsgBox sMsg, vbInformation, "Cannot set AutoNumber"
    Else
        Set rs = db.OpenRecordset("Select Max(sessionID) as SeID from SESSION")
        vMaxID = rs(0)
        If IsNull(vMaxID) Then vMaxID = 0
        If vMaxID >= lNum Then
            sMsg = "Supply a larger number. """ & sTable & "." & _
                sFieldName & """ already contains the value " & vMaxID
            MsgBox sMsg, vbInformation, "Too low."
        Else
            ' Insert and delete the record.
            sSQL = "INSERT INTO " & sTable & " ([" & sFieldName & "]) SELECT " & lNum & " AS lNum;"
            db.Execute sSQL, dbFailOnError
            sSQL = "DELETE FROM " & sTable & " WHERE " & sFieldName & " = " & lNum & ";"
            db.Execute sSQL, dbFailOnError
        End If
    End If
Exit_SetAutoNumber:
    Exit Sub

Err_SetAutoNumber:
    MsgBox "Error " & Err.Number & ": " & Err.Description, , "SetAutoNumber()"
    Resume Exit_SetAutoNumber
End Sub


Sub setupShow()
Dim rsInfo As Recordset

Set rsInfo = db.OpenRecordset("Select * from SHOWS where ShowID=" & ShowID)

'Set rsInfo = db.OpenRecordset("Select * from SHOWS where #" & Now() & "# between ShowStartDate and ShowEndDate")
'ShowID = rsInfo("ShowID")
'ShowID = 7
'lblShowName.Caption = rsInfo("ShowName")
'lblShowName.Caption = rsInfo("ShowName")
'lblStartShow.Caption = rsInfo("ShowStartDate")
'lblStopShow.Caption = rsInfo("ShowEndDate")
'lblTitles(5).Caption = "Pricing for " & Format(Now(), "MMM DD, YYYY")
'Set rsInfo = db.OpenRecordset("Select TICKET_CATEGORIES.TicketCategory,SHOW_TICKET_PRICES.* from TICKET_CATEGORIES,SHOW_TICKET_PRICES Where TICKET_CATEGORIES.TicketCategoryID=SHOW_TICKET_PRICES.TicketCategoryID and SHOW_TICKET_PRICES.ShowID=" & ShowID & " and (ShowDateStart between #" & Format(Now(), "mm/dd/yyyy") & " 12:00:00 AM# and #" & Format(Now(), "mm/dd/yyyy") & " 11:59:59 PM#) order by SHOW_TICKET_PRICES.ListSequence")
'With GridPrices
'    .Clear
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


Private Sub AN_Click()
If AN.Value Then
   bAN = True
Else
    bAN = False
End If
End Sub

Private Sub CheckBackup_Click()
Dim UseBU As String
If CheckBackup.Value = 1 Then
    UseBU = MsgBox("Are you sure that you want to use the BACKUP?", vbYesNo)

    If UseBU = vbYes Then
        bUseBU = True
    Else
        bUseBU = False
    End If

End If
End Sub

Private Sub CheckDebug_Click()
If CheckDebug.Value = 1 Then
    bDebug = True
Else
    bDebug = False

End If
End Sub

Private Sub CheckEvent_Click(Value As Integer)
If CheckEvent.Value = 0 Then
    ShowID = 9
Else
    ShowID = 7
End If
setupShow
End Sub

Private Sub checkOnLine_Click(Value As Integer)
'If checkOnLine.Value = -1 Then
'    lblSuper(0).Visible = True
'    lblSuper(1).Visible = True
'    lblSuper(2).Visible = True
'    lblSuper(3).Visible = True
'    txtUserName.Visible = True
'    txtSuperPassword.Visible = True
'    txtEventID.Visible = True
'    txtEventCode.Visible = True
'Else
'    lblSuper(0).Visible = False
'    lblSuper(1).Visible = False
'    lblSuper(2).Visible = False
'    lblSuper(3).Visible = False
'    txtUserName.Visible = False
'    txtSuperPassword.Visible = False
'    txtEventID.Visible = False
'    txtEventCode.Visible = False
'End If
End Sub

Private Sub cmdAddEdit_Click()
frmSetUp.Show
End Sub

Private Sub chkConvenienceFee_Click()
If chkConvenienceFee.Value = 0 Then
   txtConvenienceFee.text = "0.00"
End If
End Sub

Private Sub cmdBatch_Click()

frmTestTicketPrint.Show 1
End Sub

Private Sub cmdBegin_Click()
Dim strTestEvent As String
Dim rs As Recordset
Dim rsTicketstock As Recordset
bBatch = False
bNoPrintBypass = True
If CheckBackup.Value = 1 Then
    bUseBU = True
Else
    bUseBU = False
End If
If Check_AOP.Value = 1 Then
    UseAOP = True
Else
    UseAOP = False
End If
If CheckAOPBarcode.Value = 1 Then
    AOPBarcode = True
Else
    AOPBarcode = False
End If

If ComboEvents.ListIndex < 0 Then
    MsgBox "Please select EVENT!"
    Exit Sub
End If
'If CheckDebit.Value = 1 Then
'    NotAllowDebit = False
'Else
    NotAllowDebit = True
'End If
If OptionPrinter(0).Value = True Then
    TicketPrinter = 1
Else
    TicketPrinter = 1
End If
If chkZipCode.Value = 1 Then
    GetZip = True
Else
    GetZip = False
    ZipCode = ""
End If
If CheckEMail.Value = 1 Then
    bGetEMail = True
Else
    bGetEMail = False
    EMail = "NA"
End If
If CheckReceipt.Value = 1 Then
    bReceipt = True
Else
    bReceipt = False
End If
If chkConvenienceFee.Value = 1 Then
    bConvenienceFee = True
    'Set rs = db.OpenRecordset("Select * from SHOWS where ShowID=" & ShowID)
    'ConvenienceFee = rs("ConvenienceFee")
    'ConvenienceFeeType = rs("ConvenienceFeeType")
    ConvenienceFee = txtConvenienceFee.text
    If OptionCFee(0).Value Then
        ConvenienceFeeType = 1
    Else
        ConvenienceFeeType = 2
    End If

Else
    bConvenienceFee = False
End If
iMaxTix = Int(txtSecurity.text)

strPageImageCC = ""
IP = txtIP.text
SessionCounter = txtSessionCounter.text

If Int(txtSkip.text) > 0 Then
    'skip to specified Barcode
    db.Execute "UPDATE BC2" & Format(IP, "0#") & " set UsedDataTime=#" & Now() & "# where ID < " & txtSkip.text
End If
If CheckVISAMC.Value = 1 Then
    VisaMCOnly = True
    strPageImageCC = "NoAE"
Else
    VisaMCOnly = False
End If

If CheckDiscover.Value = 1 Then
    AllowDiscover = True
Else
    AllowDiscover = False
    strPageImageCC = strPageImageCC & "NoD"
End If
    
If checkDEMO.Value = 1 Then
    bDEMO = True
    ''''REMOVE
    'bNoPrintBypass = True
Else
    bDEMO = False
    'bNoPrintBypass = False
End If

db.Execute "UPDATE INIT set IPAddress='" & IP & "',SessionCounter='" & SessionCounter & "'"     ' ,SellerCode='" & SellerCode & "',Password='" & Password & "',TestEvent='" & strTestEvent & "'"
db.Execute "UPDATE SHOW_DEFAULTS set AllowAllowDiscover = " & CheckDiscover.Value & ",VISA_MC_Only =" & CheckVISAMC.Value
Set rsTicketstock = db.OpenRecordset("Select * from TICKET_STOCK")
If rsTicketstock(0) <= 10 Then
    frmStartup.Show
Else

'If UseAOP Then
    frmFirstPageAOP.Show
'Else
'    frmFirstPage.Show
'End If

End If
Unload Me


End Sub

Private Sub cmdClear_Click()
Response = MsgBox("Are you sure that you want to clear SALES HISTORY?", vbYesNo + vbCritical)
If Response = vbYes Then
    db.Execute "Delete * from PURCHASE_INFO"
    db.Execute "Delete * from TICKET_CATEGORY_RECEIPTS"
    db.Execute "Delete * from LOG"
    db.Execute "Delete * from SESSION"
    

End If
End Sub

Private Sub cmdEOD_Click()
IP = txtIP.text
SellerCode = txtSellerCode.text
Password = txtPassword.text

frmEOD.Show
Unload Me
End Sub

Private Sub cmdCoupon_Click()
frmemail.Show 1
End Sub

Private Sub cmdDeviceUP_Click()
If Int(txtIP.text) = "20" Then
   txtIP.text = "1"
Else
   txtIP.text = Trim(str(Int(txtIP.text) + 1))
   
End If
End Sub

Private Sub cmdExit_Click()
'Unload frmPINPAD

End
End Sub

Private Sub cmdLOG_Click()
frmLOG.Show 1

End Sub

Private Sub cmdMult_Click()
'If CheckDebit.Value = 1 Then
'    NotAllowDebit = False
'Else
'    NotAllowDebit = True
'End If
''If Len(Trim(txtTerminalID.Text)) = 0 Then
''  MsgBox "Please enter TERMINALID"
''  Exit Sub
''Else
''    TerminalID = Trim(txtTerminalID.Text)
''End If
'If Len(Trim(txtCity.text)) = 0 Then
'    MsgBox "Please enter an EVENT CITY"
'    Exit Sub
'Else
'    EventCity = txtCity.text
'End If
'IP = txtIP.text
'SellerCode = txtSellerCode.text
'Password = txtPassword.text
''If CheckEvent.Value = -1 Then
''    bTest = True
''    strTestEvent = "Y"
''Else
'    bTest = False
'    strTestEvent = "N"
''End If
'If checkDEMO.Value = 1 Then
'    bDEMO = True
'End If
'
'If checkOnLine.Value = -1 Then
'    Set rs = db.OpenRecordset("Select * from ADMINISTRATOR where username='" & LCase(txtUserName.text) & "' and password='" & LCase(txtSuperPassword.text) & "'")
'    If rs.EOF Then
'        MsgBox "Not a valid User Name and/or Password"
'        Exit Sub
'    Else
'        If Len(Trim(txtEventCode.text)) = 0 Or Len(Trim(txtEventID.text)) = 0 Then
'            MsgBox "Please enter BOTH and Event Code and Event ID"
'            Exit Sub
'        Else
'            bOffLine = True
'            EventID = txtEventID.text & Space(10 - Len(txtEventID.text))
'            EventCode = txtEventCode.text
'        End If
'    End If
'Else
'    bOffLine = False
'End If
'If CheckVISAMC.Value = -1 Then
'    VisaMCOnly = True
'Else
'    VisaMCOnly = False
'End If
'
'db.Execute "UPDATE INIT set IPAddress='" & IP & "',SellerCode='" & SellerCode & "',Password='" & Password & "',TestEvent='" & strTestEvent & "'"
'
'bMultiple = True
'Load frmPINPAD
'
'frmWelcomeMult.Show
'Unload Me
End Sub


Private Sub cmdSetUpNet_Click()
frmNetworkPath.Show
End Sub

Private Sub cmdSubmit_Click()
    SellerCode = Trim(txtSellerCode.text)
    Password = Trim(txtPassword.text)
    IP = Trim(txtIP.text)
    Winsock1.RemoteHost = IP
    Winsock1.RemotePort = 1130
    strCommand = "Connect"
    Winsock1.Connect

End Sub

Private Sub cmdTestInterface_Click()
IP = txtIP.text
'SellerCode = txtSellerCode.Text
'Password = txtPassword.Text
'If CheckEvent.Value = ssCBChecked Then
'    bTest = True
'Else
'    bTest = False
'End If
frmTestInterface.Show

End Sub

Private Sub cmdTestMod_Click()
For i = 1 To 18

MsgBox i Mod 8

Next
End Sub

Private Sub cmdTestSwipe_Click()
'Dim lresult As Long
'Dim output As String
'USBHID2.PortOpen = True
'
'lresult = USBHID2.USBSwipe_Command("02", output)
'MsgBox ("Result=" & CStr(lresult) & ",Data=" & USBHID2.SwipeCommandOutput)
'USBHID2.PortOpen = False

End Sub

Private Sub cmdSecurity_Click()
If Int(txtSecurity.text) = "9" Then
   txtSecurity.text = "1"
Else
   txtSecurity.text = Trim(str(Int(txtSecurity.text) + 1))
   
End If
End Sub

Private Sub cmdSessionCounter_Click()
If Int(txtSessionCounter.text) = 9 Then
   txtSessionCounter.text = "1"
Else
   txtSessionCounter.text = Trim(str(Int(txtSessionCounter.text) + 1))
   
End If

End Sub

Private Sub cmdSetuoEvent_Click()

Dim errcode As Long
Dim url As String
Dim localFileName As String

'url = "http://207.171.1.163/SuperAppUpdate.exe"
url = "http://www.cgc-services.com/TicketmationUpdate/SuperAppUpdate.exe"
localFileName = App.path & "\SuperApp" & Format(Now, "mmddyy") & ".exe"

errcode = URLDownloadToFile(0, url, localFileName, 0, 0)
If errcode = 0 Then
    MsgBox "Download ok"
    Name App.path & "\SuperApp.exe" As App.path & "\SuperAppOld" & Format(Now, "mmddyy") & ".exe"
    FileCopy App.path & "\SuperApp" & Format(Now, "mmddyy") & ".exe", App.path & "\SuperApp.exe"
    MsgBox "SUPERAPP Update Complete"
Else
    MsgBox "Error while downloading"
End If

End Sub

Private Sub cmdSkip_Click()
If Int(txtSkip.text) = "11000" Then
   txtSkip.text = "100"
Else
   txtSkip.text = Trim(str(Int(txtSkip.text) + 100))
   
End If

End Sub

Private Sub cmdTicketStock_Click()
frmTicketStock.Show 1
End Sub

Private Sub ComboCities_Change()

End Sub

Private Sub ComboEventCode_Click()
EventType = ComboEventCode.text
strCommand = "SEVT"
Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:" & Trim(txtCity.text) & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"

End Sub


Private Sub ComboEvents_Click()
Dim rsInfo As Recordset
rs.MoveFirst
rs.Move ComboEvents.ListIndex
ShowID = rs("ShowID")
SQLShowID = rs("SQLShowID")
EventType = rs("EventType")
EventCity = rs("EventCity")
ShowImage = rs("ShowImage")
EventReference = rs("EventReferenceName")
'setupShow
IP = Trim(txtIP.text)

Set rsInfo = db.OpenRecordset("Select Show_Defaults.*,Shows.ConvenienceFee,Shows.ConvenienceFeeType from Shows,Show_Defaults where Shows.showid=Show_Defaults.ShowID and Shows.ShowID=" & ShowID)
If Not rsInfo.EOF Then
chkConvenienceFee.Value = rsInfo(3)
CheckVISAMC.Value = rsInfo(2)
CheckDiscover.Value = rsInfo(1)
ConvenienceFee = rsInfo("ConvenienceFee")
txtConvenienceFee.text = ConvenienceFee
If rsInfo("ConvenienceFeeType") = 1 Then
    OptionCFee(0).Value = True
Else
    OptionCFee(1).Value = True
End If
End If
OptionPrinter(1).Value = True

End Sub


Private Sub Command1_Click()
' Call SetAutoNumber("SESSION", 312345)
i = 0
Do Until i > 1
i = 0
Loop
End Sub

Private Sub Form_Load()
'OpenDataBase
On Error GoTo errMsg
Set rs = db.OpenRecordset("Select * from Show_Defaults")
bTest = False
bDebug = False
bServer = True
bMultiple = False
bNoPorts = False
bDEMO = False
bAN = True

Set rs = db.OpenRecordset("Select * from INIT")

''txtSellerCode.Text = rs("SellerCode")
''txtPassword.Text = rs("Password")
txtIP.text = rs("IPAddress")
txtSessionCounter.text = rs("SessionCounter")

Set rs = db.OpenRecordset("Select * from SHOWS where ShowEndDate >=#" & Now() & "#  order by ShowStartDate")
If rs.EOF Then
    MsgBox "There are no current events in the database"
    ComboEvents.text = "NO EVENTS AVAILABLE"
    ComboEvents.Enabled = False
Else
    Do Until rs.EOF
        ComboEvents.AddItem rs("EventReferenceName")
        rs.MoveNext
    Loop
    If ComboEvents.ListCount = 2 Then
        ComboEvents.ListIndex = 1
    Else
        ComboEvents.ListIndex = 0
    End If
    OptionPrinter(0).Value = True
    'Use AOP BARCODE
'    CheckAOPBarcode.Value = 1
'        lblBarcode.Visible = True
'        txtSkip.Visible = True
'        cmdSkip.Visible = True
'        lblSession.Visible = False
'        txtSessionCounter.Visible = False
'        cmdSessionCounter.Visible = False

'    Use Ticketmation BARCODE
    CheckAOPBarcode.Value = 0
        lblBarcode.Visible = False
        txtSkip.Visible = False
        cmdSkip.Visible = False
        lblSession.Visible = True
        txtSessionCounter.Visible = True
        cmdSessionCounter.Visible = True
    'CheckVISAMC.Value = 0
    CheckReceipt.Value = 0
    If iMaxTix = 0 Then
        txtSecurity.text = "6"
    Else
        txtSecurity.text = CStr(iMaxTix)
    End If
End If
Exit Sub
errMsg:
MsgBox Err.Description
MsgBox rs.EOF
End Sub


Private Sub SSCommand1_Click()
frmBatch.Show 1


End Sub


Private Sub lblComputername_Click()

End Sub


Private Sub Winsock1_Connect()

    strCommand = "LOGI"
    Winsock1.SendData "LOGI " & SellerCode & "    " & Password & "                                  " & Trim(EventCity) & "                      @#end#@"

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim strData2 As String
Dim header(5) As String
Select Case strCommand
Case "LOGI"
    strCommand = "SEVT"
    Winsock1.GetData strData
    If bDebug Then MsgBox "LOGI " & strData
    Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:" & EventCity & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"

    'logged in - get Event Types for this city
'    strCommand = "CATE"
'    Winsock1.GetData strData
'    If bDebug Then MsgBox "LOGI " & strData
'    Winsock1.SendData "CATE " & SellerCode & "    " & Password & "    " & Chr(10) & "City:" & Trim(EventCity) & Chr(10) & "@#end#@"
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


Private Sub USBHID2_CardDataChanged()
'MsgBox "HERE"
'USBHID2.PortOpen = False
End Sub


Private Sub VScroll1_Change()

End Sub


Private Sub SSCommand2_Click()

End Sub


