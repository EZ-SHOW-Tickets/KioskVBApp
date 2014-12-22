VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLOG 
   Caption         =   "Log"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdClose 
      Height          =   720
      Left            =   8055
      TabIndex        =   1
      Top             =   5325
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1270
      _Version        =   196609
      Caption         =   "Close"
   End
   Begin MSFlexGridLib.MSFlexGrid GridLog 
      Height          =   5055
      Left            =   555
      TabIndex        =   0
      Top             =   210
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()

End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub Form_Load()
Dim rsSessions As Recordset
Dim rsLog As Recordset


Set rsSessions = db.OpenRecordset("Select SessionID from SESSION order by SessionID DESC")
If rsSessions.EOF Then
     With GridLog
        .Clear
        .Width = 9500
        .Rows = 2
        .ColWidth(0) = 500
        .ColWidth(1) = 3800
        .ColWidth(2) = 5000
        .Row = 0
        .Col = 0
        .text = ""
        .Col = 1
        .text = ""
        .Col = 2
        .text = "NONE"
       End With
Else
    With GridLog
        .Clear
        .Width = 9500
        .Rows = 2
        .ColWidth(0) = 500
        .ColWidth(1) = 3800
        .ColWidth(2) = 5000
        .Row = 0
        .Col = 0
        .text = "ID"
        .Col = 1
        .text = "Date/Time"
        .Col = 2
        .text = "INFORMATION"
        i = 0
        'Show Last 3 transactions
        For j = 1 To 3
                Set rsLog = db.OpenRecordset("Select LogID,LogTime,ClickFrom from LOG where SessionID=" & rsSessions(0) & " order by LogID")
                Do Until rsLog.EOF
                    .Rows = .Rows + 1
                    .Row = i + 1
                    .Col = 0
                    .text = rsSessions(0)
                    .Col = 1
                    .text = rsLog(1)
                    .Col = 2
                    .text = rsLog(2)
                    i = i + 1
                    rsLog.MoveNext
                Loop
                'Insert Blanck Row
                    .Rows = .Rows + 1
                    .Row = i + 1
                    .Col = 0
                    .CellBackColor = &H8000000F
                    .text = "  "
                    .Col = 1
                    .CellBackColor = &H8000000F
                    .text = "  "
                    .Col = 2
                    .CellBackColor = &H8000000F
                    .text = "  "
                    i = i + 1
            rsSessions.MoveNext
        Next j
    End With
End If
End Sub


