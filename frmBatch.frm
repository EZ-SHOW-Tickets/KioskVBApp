VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBatch 
   Caption         =   "Batch Upload"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdDelete 
      Height          =   810
      Left            =   2625
      TabIndex        =   4
      Top             =   4230
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1429
      _Version        =   196609
      Caption         =   "DELETE"
   End
   Begin MSFlexGridLib.MSFlexGrid gridBatch 
      Height          =   3675
      Left            =   420
      TabIndex        =   2
      Top             =   450
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin Threed.SSCommand cmdUpload 
      Height          =   840
      Left            =   195
      TabIndex        =   1
      Top             =   4215
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1482
      _Version        =   196609
      Caption         =   "UPLOAD"
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   675
      Left            =   5700
      TabIndex        =   0
      Top             =   4395
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1191
      _Version        =   196609
      Caption         =   "EXIT"
   End
   Begin VB.Label LabelWebState 
      Alignment       =   2  'Center
      Caption         =   "The Internet Connection is DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label LabelNone 
      Caption         =   "No batches need to be uploaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   870
      TabIndex        =   3
      Top             =   1395
      Width           =   4920
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBatch As Recordset
Dim rsBatchDetails As Recordset

Sub PopulateGrid()
Set rsBatch = db.OpenRecordset("Select * from BATCH where BatchUpload=No")
gridBatch.Clear
If rsBatch.EOF Then
    gridBatch.Visible = False
Else
    With gridBatch
        .Width = 9500
        .Rows = 2
        .ColWidth(0) = 300
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .ColWidth(4) = 2000
        .Row = 0
        .Col = 0
        .text = "ID"
        .Col = 1
        .text = "Track1"
        .Col = 2
        .text = "Track2"
        .Col = 3
        .text = "Price"
        .Col = 4
        .text = "Date"
        i = 0
        Do Until rsBatch.EOF
            .Rows = .Rows + 1
            .Row = i + 1
            .Col = 0
            .text = rsBatch(0)
            .Col = 1
            .text = rsBatch(1)
            .Col = 2
            .text = rsBatch(2)
            .Col = 3
            .text = rsBatch(3)
            .Col = 4
            .text = rsBatch(4)
            i = i + 1
           
            rsBatch.MoveNext
        Loop
    End With
End If

End Sub


Private Sub cmdDelete_Click()
Dim bRow As Integer
Dim eRow As Integer

With gridBatch
    bRow = .Row
    eRow = .RowSel
    .Col = 0
    For i = bRow To eRow
      .Row = i
      db.Execute "Delete * from BATCH where BatchID=" & .text
    Next i
End With
PopulateGrid

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpload_Click()
Dim i As Integer
Dim j As Integer
bBatch = True
'******USE DATABASE
Set rsBatch = db.OpenRecordset("Select * from BATCH where BatchUpload=No")
If rsBatch.EOF Then
    gridBatch.Visible = False
Else
    Do Until rsBatch.EOF
        Set rsBatchDetails = db.OpenRecordset("select * from BATCH_DETAILS where BatchID=" & rsBatch(0))
        i = 0
        ReDim TicketTypes(5, 0)
        ReDim TotalTicketSold(0)
        Do Until rsBatchDetails.EOF
            ReDim Preserve TicketTypes(5, i)
            ReDim Preserve TotalTicketSold(i)
            TicketTypes(1, i) = rsBatchDetails(1)
            TotalTicketSold(i) = rsBatchDetails(2)
            TotalPrice = rsBatch(3)
            i = i + 1
            rsBatchDetails.MoveNext
        Loop
        Track1 = rsBatch(1)
        AccountNumber = rsBatch(2)
        CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
        frmWebBrowser.Show 1
        'ANTransactionCode = "0"
        'Remove from batch
        If Left(ANTransactionCode, 1) <> "-" Then
            db.Execute "UPDATE BATCH set BatchUpload = Yes,Track1='" & ANTransactionCode & "',Track2='" & Now() & "' Where BatchID=" & rsBatch(0)
            LabelWebState.Visible = False
        Else
            LabelWebState.Visible = True
        End If

        rsBatch.MoveNext
    Loop
End If

'*************USE GRID
'LOOP through grid

'With gridBatch
'    j = 1
'    .Row = j
'    .Col = 0
'    Do Until Len(.Text) = 0
'    'For i = 1 To .Rows
'    '    .Row = i
'        .Col = 1
'        Track1 = .Text
'        .Col = 2
'        AccountNumber = .Text
'        .Col = 3
'        TotalPrice = CInt(.Text)
'        .Col = 0
'        Set rsBatchDetails = db.OpenRecordset("select * from BATCH_DETAILS where BatchID=" & .Text)
'        i = 0
'        ReDim TicketTypes(1, 0)
'        ReDim TotalTicketSold(0)
'        i = -1
'        Do Until rsBatchDetails.EOF
'            i = i + 1
'            ReDim Preserve TicketTypes(1, i)
'            ReDim Preserve TotalTicketSold(i)
'            TicketTypes(1, i) = rsBatchDetails(1)
'            TotalTicketSold(i) = rsBatchDetails(2)
'            rsBatchDetails.MoveNext
'        Loop
'        If InStr(Track1, "^") = 0 Then
'            CardName = "NA"
'        Else
'            CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
'        End If
'        SwipTime = Now()
'        frmWebBrowser.Show 1
'        If ANTransactionCode <> "-1" Then
'            db.Execute "UPDATE BATCH set BatchUpload = Yes Where BatchID=" & .Text
'            LabelWebState.Visible = False
'        Else
'            LabelWebState.Visible = True
'        End If
'        j = j + 1
'        .Row = j
'        .Col = 0
'    Loop
'End With
bBatch = False
PopulateGrid
End Sub

Private Sub Form_Load()
Dim i As Integer
frmBatch.Width = 10000

PopulateGrid

End Sub


