VERSION 5.00
Begin VB.Form frmKluge 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   90
      Top             =   75
   End
End
Attribute VB_Name = "frmKluge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For i = 1 To 10
     DoEvents
Next i
Call LogClick("In Kluge", "Start Timer")

'frmSpecifyTickets.Show
'Unload Me
Timer1.interval = 100
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Call LogClick("In Kluge", "stop Timer")
Timer1.interval = 0
Timer1.Enabled = False
Call LogClick("In Kluge", "Load frmSpecifyTickets")
frmSpecifyTickets.Show
Call LogClick("In Kluge", "Unload")
Unload Me
End Sub


