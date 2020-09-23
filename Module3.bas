Attribute VB_Name = "Module3"
Function MainLoop()
Do While Running = True
GetKeys
UpdateMouse
SetNewRally
ChkWorkerSelect
ChkWorkerCollision
CHKBuildUnit
CHKConstruct
CHKGoldReturn
CHKMouseBuild
Mine
UnitGoto
BLT
FPSCount = FPSCount + 1
DoEvents
Loop
CleanUp
End Function

Function UpdateMouse()
GetCursorPos MousePos
End Function

Function ChkWorkerSelect()
For I = 1 To 20 Step 1
With Worker(I)
    If MousePos.X > Mouse.LastX And MousePos.Y > Mouse.LastY Then
        If .X < MousePos.X And .X > Mouse.LastX And .Y < MousePos.Y And .Y > Mouse.LastY And Mouse.GotInfo = True Then
            .Selected = True
        End If
    ElseIf MousePos.X > Mouse.LastX And MousePos.Y < Mouse.LastY Then
        If .X < MousePos.X And .X > Mouse.LastX And .Y > MousePos.Y And .Y < Mouse.LastY And Mouse.GotInfo = True Then
            .Selected = True
        End If
    ElseIf MousePos.X < Mouse.LastX And MousePos.Y > Mouse.LastY Then
        If .X > MousePos.X And .X < Mouse.LastX And .Y < MousePos.Y And .Y > Mouse.LastY And Mouse.GotInfo = True Then
            .Selected = True
        End If
    ElseIf MousePos.X < Mouse.LastX And MousePos.Y < Mouse.LastY Then
        If .X > MousePos.X And .X < Mouse.LastX And .Y > MousePos.Y And .Y < Mouse.LastY And Mouse.GotInfo = True Then
            .Selected = True
        End If
    End If
End With
Next
End Function

Function UnSelectWorkers()
If MousePos.X < 810 And MousePos.Y < 630 Then
    For I = 1 To 20 Step 1
        With Worker(I)
            If .Selected = True Then
                .Selected = False
            End If
        End With
    Next
End If
End Function

Function UnitGoto()
For I = 1 To 20 Step 1
With Worker(I)
    If .X <> .NX Then
        If .X < .NX Then
            .X = .X + 1
        Else
            .X = .X - 1
        End If
    End If
    If .Y <> .NY Then
        If .Y < .NY Then
            .Y = .Y + 1
        Else
            .Y = .Y - 1
        End If
    End If
    
End With
Next
End Function

Function ChkWorkerCollision()
For w1 = 1 To 20 Step 1
    For w2 = 1 To 20 Step 1
        If Worker(w1).Mining <> 1 Then
            If Worker(w1).NX = Worker(w2).NX And Worker(w1).NY = Worker(w2).NY And w1 <> w2 Then
                Worker(w1).NX = Worker(w1).NX - 64
            End If
        End If
    Next
Next
End Function

Function CHKMine(I As Long)
If MousePos.X > GoldMine(1).X And MousePos.X < GoldMine(1).X + 128 And MousePos.Y > GoldMine(1).Y And MousePos.Y < GoldMine(1).Y + 128 Then
    If Worker(I).Selected = True Then
        Worker(I).NX = GoldMine(1).X + 64
        Worker(I).NY = GoldMine(1).Y + 64
        Worker(I).Mining = 1
    End If
End If
End Function

Function Mine()
For I = 1 To 20 Step 1
    With Worker(I)
        If .Mining = 2 Then
            .NX = TownHall(.TownHall).X + 64
            .NY = TownHall(.TownHall).Y + 64
            If .X = .NX And .Y = .NY Then
                .Mining = 1
                Player.Gold = Player.Gold + 10
            End If
        ElseIf .Mining = 1 Then
            .NX = GoldMine(1).X + 64
            .NY = GoldMine(1).Y + 64
            If .X = .NX And .Y = .NY Then
                .Mining = 2
            End If
        End If
    End With
Next
End Function

Function WorkerGotoNext()
For I = 1 To 20 Step 1
    If Worker(I).Selected = True Then
        Worker(I).NX = MousePos.X
        Worker(I).NY = MousePos.Y
        If Worker(I).Mining = 2 Or Worker(I).Mining = 3 Then
            Worker(I).Mining = 3
        Else
            Worker(I).Mining = 0
    End If
        CHKMine (I)
    End If
Next
End Function

Function SetNewRally()
For I = 1 To 2 Step 1
    If MousePos.X > 885 And MousePos.X < 930 And MousePos.Y > 650 And MousePos.Y < 695 And Mouse.Button = "Left" And TownHall(I).RallySet = True And TownHall(I).Selected = True Then
        TownHall(I).RallySet = False
    ElseIf MousePos.X < 830 And MousePos.Y < 650 And Mouse.Button = "Right" And TownHall(I).RallySet = False Then
        TownHall(I).RallySet = True
    End If
    If TownHall(I).RallySet = False Then
        TownHall(I).RallyX = MousePos.X
        TownHall(I).RallyY = MousePos.Y
    End If
Next
End Function
