Attribute VB_Name = "Module2"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Running As Boolean
Public FPS As Long
Public FPSCount As Long

Public Type PointAPI
    X As Long
    Y As Long
End Type
Global MousePos As PointAPI

Public Type Mouse
    X As Long
    Y As Long
    LastX As Long
    LastY As Long
    Button As String
    GotInfo As Boolean
End Type
Global Mouse As Mouse

Public Type GoldMine
    Gold As Long
    X As Long
    Y As Long
    Depleted As Boolean
End Type
Global GoldMine(1 To 5) As GoldMine

Public Type TownHall
    X As Long
    Y As Long
    RallyX As Long
    RallyY As Long
    RallySet As Boolean
    Health As Long
    Building As Boolean
    KeepUnits As KeepUnits
    Selected As Boolean
    active As Boolean
    prog As Long
    start As Boolean
End Type
Global TownHall(1 To 2)  As TownHall

Public Type Barracks
    X As Long
    Y As Long
    RallyX As Long
    RallyY As Long
    RallySet As Boolean
    Health As Long
    Building As Boolean
    Selected As Boolean
    active As Boolean
    prog As Long
    start As Boolean
    Nextup As Long
End Type
Global Barracks(1 To 5) As Barracks



Public Type Worker
    X As Long
    Y As Long
    Health As Long
    Vis As Boolean
    Selected As Boolean
    NX As Long
    NY As Long
    Mining As Long
    Nextup As Long
    Construct As Long
    TownHall As Long
    GoldReq As Long
    prog As Long
    start As Boolean
End Type
Global Worker(1 To 20) As Worker

Public Type Player
    Gold As Long
End Type
Global Player As Player

    
Public Enum KeepUnits
    Peasent = 1
    King = 2
End Enum



Function Startup()
Running = True
FPS = 0
FPSCount = 0

With GoldMine(1)
    .X = 100
    .Y = 100
    .Gold = 1000
    .Depleted = False
End With

With TownHall(1)
    .X = 500
    .Y = 500
    .Health = 100
    .KeepUnits = Peasent
    .Building = False
    .Selected = False
    .active = True
    .prog = 0
    .start = False
    .RallyX = 500
    .RallyY = 500
    .RallySet = True
End With

For I = 1 To 5 Step 1
    With Barracks(I)
        .active = False
        .Building = False
        .Health = 100
        .Nextup = 1
        .prog = 0
        .Selected = False
        .start = False
        .RallySet = True
    End With
Next

With TownHall(2)
    .X = 1024
    .active = False
End With

For I = 1 To 20 Step 1
    With Worker(I)
        .Health = 3
        .Selected = False
        .Mining = 0
        .Construct = 0
        .TownHall = 1
        .GoldReq = 50
    End With
Next

With Worker(1)
    .X = 300
    .Y = 300
    .NY = 300
    .NX = 300
    .Vis = True
    .Nextup = 3
End With

With Worker(2)
    .X = 400
    .Y = 300
    .NX = 400
    .NY = 300
    .Vis = True
End With

With Player
    .Gold = 50
End With

End Function


'Mining
'0 = not
'1 = going to mine
'2 = returning gold
'3 = holding gold

'Construction
'0 = not
'1 = select where to build \
'2 = go to construct        }- townhall
'3 = construct             /
'4 = select where to build \
'5 = goto construct         }- barracks
'6 = construct             /
