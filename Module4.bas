Attribute VB_Name = "Module4"
Function CHKBuildingSelect()
For I = 1 To 2 Step 1
    With TownHall(I)
        If MousePos.X > .X And MousePos.X < .X + 128 And MousePos.Y > .Y _
        And MousePos.Y < .Y + 128 Or MousePos.X > 810 And MousePos.Y > 630 And .Selected = True And .active = True Then
            .Selected = True
        Else
            .Selected = False
        End If
    End With
Next
For I = 1 To 5 Step 1
    With Barracks(I)
        If MousePos.X > .X And MousePos.X < .X + 128 And MousePos.Y > .Y _
        And MousePos.Y < .Y + 128 Or MousePos.X > 810 And MousePos.Y > 630 And .Selected = True And .active = True Then
            .Selected = True
        Else
            .Selected = False
        End If
    End With
Next
End Function

Function CHKBuildUnit()
For I = 1 To 2 Step 1
With MousePos
    If .X > 830 And .Y > 650 And .X < 875 And .Y < 695 And TownHall(I).Selected = True And Mouse.Button = "Left" Then
            If Player.Gold > Worker(1).GoldReq - 1 Then
                TownHall(I).start = True
            End If
    End If
End With

If TownHall(I).start = True And TownHall(I).prog < 201 Then
    TownHall(I).prog = TownHall(I).prog + 1
End If

If Worker(1).Nextup < 20 Then

If TownHall(I).prog = 200 Then
    TownHall(I).start = False
    TownHall(I).prog = 0
    With Worker(Worker(1).Nextup)
        .Vis = True
        .X = TownHall(I).X + 64
        .Y = TownHall(I).Y + 64
        .NX = TownHall(I).RallyX
        .NY = TownHall(I).RallyY
        .Selected = False
        .Mining = 0
        .Construct = 0
        .Health = 3
    End With
    Worker(1).Nextup = Worker(1).Nextup + 1
    Player.Gold = Player.Gold - Worker(1).GoldReq
End If
End If
Next
End Function

Function CHKConstruct()
For I = 1 To 20 Step 1
    With MousePos
        If .X > 830 And .Y > 650 And .X < 875 And .Y < 695 And Worker(I).Selected = True And Worker(I).Construct = 0 And Mouse.Button = "Left" Then
            Worker(I).Construct = 1
            If Worker(I).Mining <> 2 Then
                Worker(I).Mining = 0
            Else                            ''townhall
                Worker(I).Mining = 3
            End If
        End If
        
        If .X > 885 And .Y > 650 And .X < 935 And .Y < 695 And Worker(I).Selected = True And Worker(I).Construct = 0 And Mouse.Button = "Left" Then
            Worker(I).Construct = 4
            If Worker(I).Mining <> 2 Then
                Worker(I).Mining = 0
            Else                            ''barracks
                Worker(I).Mining = 3
            End If
        End If
    End With
    
    If Worker(I).Construct = 1 Then
        TownHall(2).X = MousePos.X
        TownHall(2).Y = MousePos.Y
    End If
    If Worker(I).Construct = 4 Then
        Barracks(Barracks(1).Nextup).X = MousePos.X
        Barracks(Barracks(1).Nextup).Y = MousePos.Y
    End If
        
    If Worker(I).Construct = 2 Then
        Worker(I).NX = TownHall(2).X
        Worker(I).NY = TownHall(2).Y
        Worker(I).Construct = 3
        Worker(I).Selected = True
    End If
    If Worker(I).Construct = 5 Then
        Worker(I).NX = Barracks(Barracks(1).Nextup).X
        Worker(I).NY = Barracks(Barracks(1).Nextup).Y
        Worker(I).Construct = 6
        Worker(I).Selected = True
    End If
    
    If Worker(I).Construct = 3 And Worker(I).X = Worker(I).NX And Worker(I).Y = Worker(I).NY Then
        Worker(I).NX = Worker(I).X
        Worker(I).NY = Worker(I).Y
        Worker(I).prog = Worker(I).prog + 1
        If Worker(I).prog = 200 Then
            Worker(I).prog = 0
            Worker(I).start = False
            Worker(I).Construct = 0
            TownHall(2).prog = 0
            TownHall(2).start = False
            TownHall(2).active = True
            CHKMineDist
        End If
    End If
    
    If Worker(I).Construct = 6 And Worker(I).X = Worker(I).NX And Worker(I).Y = Worker(I).NY Then
        Worker(I).NX = Worker(I).X
        Worker(I).NY = Worker(I).Y
        Worker(I).prog = Worker(I).prog + 1
        If Worker(I).prog = 200 Then
            Worker(I).prog = 0
            Worker(I).start = False
            Worker(I).Construct = 0
            Barracks(Barracks(1).Nextup).active = True
            Barracks(1).Nextup = Barracks(1).Nextup + 1
        End If
    End If
Next



End Function

Function CHKMouseBuild()
For I = 1 To 20 Step 1
    With Worker(I)
        If .Construct = 1 And Mouse.Button = "Left" And MousePos.X < 830 And MousePos.Y < 650 Then
            .Construct = 2
            .NX = MousePos.X
            .NY = MousePos.Y
            TownHall(2).X = MousePos.X
            TownHall(2).Y = MousePos.Y
            TownHall(2).Selected = False
            TownHall(2).Building = False
        End If
        
        If .Construct = 4 And Mouse.Button = "Left" And MousePos.X < 830 And MousePos.Y < 650 Then
            .Construct = 5
            .NX = MousePos.X
            .NY = MousePos.Y
            Barracks(Barracks(1).Nextup).X = MousePos.X
            Barracks(Barracks(1).Nextup).Y = MousePos.Y
        End If
    End With
Next
End Function

Function CHKGoldReturn()
If MousePos.X > 945 And MousePos.Y > 650 And MousePos.X < 990 And MousePos.Y < 695 And Mouse.Button = "Left" Then
    For I = 1 To 5 Step 1
        If Worker(I).Selected = True Then
            Worker(I).Mining = 2
        End If
    Next
End If
End Function

Function CHKMineDist()
Dim th1d As Long
Dim th2d As Long
For I = 1 To 5 Step 1 '********************change to 20 for all guys !!***************
If TownHall(2).active = True Then
    If TownHall(1).X > GoldMine(1).X Then
        th1d = TownHall(1).X - GoldMine(1).X
    Else
        th1d = GoldMine(1).X - TownHall(1).X
    End If
    If TownHall(2).X > GoldMine(1).X Then
        th2d = TownHall(2).X - GoldMine(1).X
    Else
        th2d = GoldMine(1).X - TownHall(2).X
    End If
    If th1d > th2d Then
        Worker(I).TownHall = 2
    Else
        Worker(I).TownHall = 1
    End If
End If
Next
End Function


