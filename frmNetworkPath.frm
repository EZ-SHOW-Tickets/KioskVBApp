VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNetworkPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Path"
   ClientHeight    =   3270
   ClientLeft      =   4200
   ClientTop       =   4155
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNetworkPath.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5220
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   3975
      TabIndex        =   4
      Top             =   1110
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog cmdgNetworkPath 
      Left            =   45
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3975
      TabIndex        =   2
      Top             =   1590
      Width           =   1185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   435
      Left            =   3990
      TabIndex        =   1
      Top             =   480
      Width           =   1170
   End
   Begin VB.Label lblPathTitle 
      Caption         =   "Path of Database:"
      Height          =   240
      Left            =   105
      TabIndex        =   5
      Top             =   975
      Width           =   1335
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1230
      Width           =   3720
   End
   Begin VB.Label lblDirections 
      Caption         =   "Locate the Drive and Path for the Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4545
   End
End
Attribute VB_Name = "frmNetworkPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.10.32 generated this copy of frmNetworkPath.frm on Tue 7/27/99 @ 1
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  mstrFileTitle                                                                         *
'******************************************************************************************

Option Explicit

'// Private variables
Private mstrNetPath     As String
Private mstrFileName    As String
Private mstrStoreKey    As String

Private Sub cmdBrowse_Click()
'On Error GoTo EH_Show

    With cmdgNetworkPath
        .ShowSave
        mstrFileName = .FileTitle
        mstrNetPath = Left$(.FileName, Len(.FileName) - Len(.FileTitle))
    End With

    lblPath.Caption = mstrNetPath

'Exit Sub
'EH_Show:
'    ShowError Err, App.EXEName & "." & TypeName(Me) & ".cmdBrowse_Click", Me
End Sub

Private Sub cmdCancel_Click()
'On Error GoTo EH_Show

    Dim strNP   As String

    'strNP = GetSetting("ISTCleanware", mstrStoreKey, "NetworkPath")
    'If Len(strNP) < 1 Then
    '    MsgBox "Select a Database to continuing."
    'Else
       Unload Me
    'End If

'Exit Sub
'EH_Show:
'    ShowError Err, App.EXEName & "." & TypeName(Me) & ".cmdCancel_Click", Me
End Sub

Private Sub cmdOK_Click()
'On Error GoTo EH_Show
    Dim strpath As String
    Dim dbLocal As Database
    Dim Response As Integer
    
    strpath = App.Path & "\WUS.mdb"

    If Len(mstrNetPath) > 0 Then
        Response = MsgBox("This is the path that your database is located: " & mstrNetPath & ". " & Chr(9) & Chr(13) & "This is critical to run the system. " & Chr(9) & Chr(13) & "You must restart iTrade for this to be updated.  iTrade will be shut down now", vbYesNo + vbCritical + vbDefaultButton1, "Database Path")
        If Response = vbYes Then
            Set dbLocal = wrkjet.OpenDataBase(strpath)
            dbLocal.Execute "UPDATE SETUP_INFO set DatabaseServer='" & mstrNetPath & "'"
'            SaveSetting "ISTCleanware", mstrStoreKey, "NetworkPath", mstrNetPath
            dbLocal.Close
            db.Close
            Set db = wrkjet.OpenDataBase(mstrNetPath & "\WUS.mdb")
            End
            'Unload Me
        End If
    Else
       MsgBox "You have NOT chosen a path. Click the Browse button to select one.", vbInformation, "Database Path"
    End If

'Exit Sub
'EH_Show:
 '   ShowError Err, App.EXEName & "." & TypeName(Me) & ".cmdOK_Click", Me
End Sub

Private Sub Form_Load()
'On Error GoTo EH_Log
    Dim rsInfo As Recordset
    Set rsInfo = db.OpenRecordset("Select * from SETUP_INFO")
    mstrNetPath = rsInfo("DatabaseServer")
    lblPath.Caption = mstrNetPath

Exit Sub
'EH_Log:
    'LogError Err, App.EXEName & "." & TypeName(Me) & ".Form_Load"
'    Resume Next '// Must continue anyway
End Sub


'
Private Sub lblDirections_Click()

End Sub


