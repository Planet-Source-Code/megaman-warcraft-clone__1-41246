VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
DoEvents
DIINIT
DDINIT
INITSurfaces
Startup
MainLoop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Mouse.GotInfo = False Then
    Mouse.Button = "Left"
    Mouse.GotInfo = True
    Mouse.LastX = MousePos.X
    Mouse.LastY = MousePos.Y
    UnSelectWorkers
    CHKBuildingSelect
End If

If Button = 2 Then
    Mouse.Button = "Right"
    WorkerGotoNext
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse.Button = "none"
Mouse.GotInfo = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
End Sub

Private Sub Timer1_Timer()
FPS = FPSCount
FPSCount = 0
End Sub
