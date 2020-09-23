Attribute VB_Name = "Module1"
Dim DX As New DirectX7

Dim DD As DirectDraw7

Dim Screen As DirectDrawSurface7    'ddsd1
Dim Buffer As DirectDrawSurface7    'ddsd2
Dim Ground As DirectDrawSurface7    'ddsd3
Dim Gold As DirectDrawSurface7      'ddsd4
Dim Keep As DirectDrawSurface7      'ddsd4
Dim Up As DirectDrawSurface7        'ddsd5
Dim Down As DirectDrawSurface7
Dim Right As DirectDrawSurface7
Dim Left As DirectDrawSurface7
Dim UpGold As DirectDrawSurface7
Dim DownGold As DirectDrawSurface7
Dim LeftGold As DirectDrawSurface7
Dim RightGold As DirectDrawSurface7
Dim Health1 As DirectDrawSurface7
Dim Health2 As DirectDrawSurface7
Dim Health3 As DirectDrawSurface7
Dim Hud As DirectDrawSurface7
Dim WButton As DirectDrawSurface7
Dim TButton As DirectDrawSurface7
Dim BButton As DirectDrawSurface7
Dim RButton As DirectDrawSurface7
Dim AButton As DirectDrawSurface7
Dim GReturn As DirectDrawSurface7
Dim GKeep As DirectDrawSurface7
Dim GBarracks As DirectDrawSurface7
Dim DBarracks As DirectDrawSurface7
Dim Rally As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2
Dim ddsd4 As DDSURFACEDESC2
Dim ddsd5 As DDSURFACEDESC2
Dim ddsd6 As DDSURFACEDESC2
Dim Caps As DDSCAPS2

Dim ButtonRect As RECT
Dim BackRect As RECT
Dim BuildingRect As RECT
Dim UnitRect As RECT

Dim DI As DirectInput
Dim DIDev As DirectInputDevice
Dim DIState As DIKEYBOARDSTATE

Function DIINIT()
Set DI = DX.DirectInputCreate
Set DIDev = DI.CreateDevice("guid_syskeyboard")
DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
DIDev.SetCooperativeLevel Form1.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
DIDev.Acquire
End Function

Function GetKeys()
DIDev.GetDeviceStateKeyboard DIState
With DIState
    If .Key(1) Then
        Running = False
    End If
End With
End Function

Function DDINIT()
On Error Resume Next
Set DD = DX.DirectDrawCreate("")
DD.SetCooperativeLevel Form1.hWnd, DDSCL_ALLOWMODEX Or DDSCL_ALLOWREBOOT Or DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
DD.SetDisplayMode 1024, 768, 16, 0, DDSDM_DEFAULT

ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ddsd1.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
ddsd1.lBackBufferCount = 1
Set Screen = DD.CreateSurface(ddsd1)

Caps.lCaps = DDSCAPS_BACKBUFFER
Set Buffer = Screen.GetAttachedSurface(Caps)
Buffer.GetSurfaceDesc ddsd2

End Function

Function INITSurfaces()
On Error Resume Next

ddsd3.lFlags = DDSD_CAPS
ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set Ground = DD.CreateSurfaceFromFile(App.Path + "\grass.bmp", ddsd3)
Set Hud = DD.CreateSurfaceFromFile(App.Path + "\hud.bmp", ddsd3)

ddsd4.lFlags = DDSD_CAPS
ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set Gold = DD.CreateSurfaceFromFile(App.Path + "\sprites\buildings\gold mine.bmp", ddsd4)
Set Keep = DD.CreateSurfaceFromFile(App.Path + "\sprites\buildings\keep.bmp", ddsd4)
Set GKeep = DD.CreateSurfaceFromFile(App.Path + "\sprites\buildings\green keep.bmp", ddsd4)
Set GBarracks = DD.CreateSurfaceFromFile(App.Path + "\sprites\buildings\barracksgreen.bmp", ddsd4)
Set DBarracks = DD.CreateSurfaceFromFile(App.Path + "\sprites\buildings\barracks.bmp", ddsd4)


ddsd5.lFlags = DDSD_CAPS
ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set Up = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\up.bmp", ddsd5)
Set Down = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\down.bmp", ddsd5)
Set Left = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\left.bmp", ddsd5)
Set Right = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\right.bmp", ddsd5)
Set UpGold = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\upgold.bmp", ddsd5)
Set DownGold = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\downgold.bmp", ddsd5)
Set LeftGold = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\leftgold.bmp", ddsd5)
Set RightGold = DD.CreateSurfaceFromFile(App.Path + "\sprites\peasent\rightgold.bmp", ddsd5)
Set Health3 = DD.CreateSurfaceFromFile(App.Path + "\sprites\health\green.bmp", ddsd5)
Set Health2 = DD.CreateSurfaceFromFile(App.Path + "\sprites\health\yellow.bmp", ddsd5)
Set Health1 = DD.CreateSurfaceFromFile(App.Path + "\sprites\health\red.bmp", ddsd5)
Set Rally = DD.CreateSurfaceFromFile(App.Path + "\sprites\flag.bmp", ddsd5)

ddsd6.lFlags = DDSD_CAPS
ddsd6.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set WButton = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\worker.bmp", ddsd6)
Set TButton = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\TownHall.bmp", ddsd6)
Set GReturn = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\gold return.bmp", ddsd6)
Set BButton = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\barracks.bmp", ddsd6)
Set RButton = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\rally.bmp", ddsd6)
Set AButton = DD.CreateSurfaceFromFile(App.Path + "\sprites\pictures\Archer.bmp", ddsd6)


DX.GetWindowRect Form1.hWnd, BackRect
BuildingRect.Bottom = 128
BuildingRect.Right = 128
ButtonRect.Right = 45
ButtonRect.Bottom = 45
UnitRect.Bottom = 64
UnitRect.Right = 64

Buffer.SetFontTransparency True
Buffer.SetForeColor &HFF00
Buffer.SetFillColor &H0


Dim Key As DDCOLORKEY
Key.low = 0
Key.high = 0

Gold.SetColorKey DDCKEY_SRCBLT, Key
Keep.SetColorKey DDCKEY_SRCBLT, Key
Down.SetColorKey DDCKEY_SRCBLT, Key
DownGold.SetColorKey DDCKEY_SRCBLT, Key
Up.SetColorKey DDCKEY_SRCBLT, Key
UpGold.SetColorKey DDCKEY_SRCBLT, Key
Right.SetColorKey DDCKEY_SRCBLT, Key
RightGold.SetColorKey DDCKEY_SRCBLT, Key
Left.SetColorKey DDCKEY_SRCBLT, Key
LeftGold.SetColorKey DDCKEY_SRCBLT, Key
Health1.SetColorKey DDCKEY_SRCBLT, Key
Health2.SetColorKey DDCKEY_SRCBLT, Key
Health3.SetColorKey DDCKEY_SRCBLT, Key
Hud.SetColorKey DDCKEY_SRCBLT, Key
WButton.SetColorKey DDCKEY_SRCBLT, Key
TButton.SetColorKey DDCKEY_SRCBLT, Key
BButton.SetColorKey DDCKEY_SRCBLT, Key
RButton.SetColorKey DDCKEY_SRCBLT, Key
AButton.SetColorKey DDCKEY_SRCBLT, Key
GKeep.SetColorKey DDCKEY_SRCBLT, Key
GBarracks.SetColorKey DDCKEY_SRCBLT, Key
DBarracks.SetColorKey DDCKEY_SRCBLT, Key
GReturn.SetColorKey DDCKEY_SRCBLT, Key
Rally.SetColorKey DDCKEY_SRCBLT, Key
End Function


Function BLT()
On Error Resume Next

Buffer.BltFast 0, 0, Ground, BackRect, DDBLTFAST_WAIT

Buffer.BltFast GoldMine(1).X, GoldMine(1).Y, Gold, BuildingRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
DrawTownHall
DrawBarracks

DrawBuildLoc
DrawWorker
DrawHealth

DrawMouseRect


Buffer.BltFast 0, 0, Hud, BackRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

DrawButtons

DrawProduction

Buffer.DrawText 0, 0, "FPS = " & FPS, False
Buffer.DrawText 0, 15, "MouseX = " & MousePos.X, False
Buffer.DrawText 0, 30, "MouseY = " & MousePos.Y, False
Buffer.DrawText 0, 45, "Number of Guys = " & Worker(1).Nextup - 1, False

Buffer.DrawText 520, 5, "Gold = " & Player.Gold, False

Screen.Flip Nothing, DDFLIP_WAIT
End Function


Function CleanUp()
On Error Resume Next
DD.RestoreDisplayMode
DD.SetCooperativeLevel Form1.hWnd, DDSCL_NORMAL
DIDev.Unacquire
Set DI = Nothing
Set DD = Nothing
Set DX = Nothing
End
End Function

Function DrawWorker()
For I = 1 To 20 Step 1
    With Worker(I)
        If .Vis = True Then
            If .Y > .NY And .X = .NX Then       '****up****     'done
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Up, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, UpGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y > .NY And .X < .NX Then   '**up&right*
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y > .NY And .X > .NX Then   '**up&Left*
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y < .NY And .X = .NX Then   '**Down***      'done
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y < .NY And .X < .NX Then   '*down&right
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y < .NY And .X > .NX Then   '*down&left*
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y = .NY And .X < .NX Then   '***right**     'done
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Right, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, RightGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y = .NY And .X > .NX Then   '****left**     'done
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Left, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, LeftGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            ElseIf .Y = .NY And .X = .NX Then   '**stopped*     'done
                If .Mining = 0 Or .Mining = 1 Then
                    Buffer.BltFast .X - 32, .Y - 32, Down, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    Buffer.BltFast .X - 32, .Y - 32, DownGold, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        End If
    End With
Next
End Function

Function DrawHealth()
For I = 1 To 20 Step 1
With Worker(I)
    If .Selected = True Then
        If .Health = 1 Then
            Buffer.BltFast .X - 32, .Y - 32, Health1, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf .Health = 2 Then
            Buffer.BltFast .X - 32, .Y - 32, Health2, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf .Health = 3 Then
            Buffer.BltFast .X - 32, .Y - 32, Health3, UnitRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If
End With
Next
End Function

Function DrawMouseRect()
If Mouse.GotInfo = True Then
    Buffer.DrawBox Mouse.LastX, Mouse.LastY, MousePos.X, MousePos.Y
End If
End Function

Function DrawButtons()
For I = 1 To 2 Step 1
    If TownHall(I).Selected = True And TownHall(I).active = True Then
        Buffer.BltFast 830, 650, WButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Buffer.BltFast 885, 650, RButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Buffer.BltFast TownHall(I).RallyX, TownHall(I).RallyY, Rally, UnitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
Next

For I = 1 To 20 Step 1
    If Worker(I).Selected = True Then
        If TownHall(2).active = False Then
            Buffer.BltFast 830, 650, TButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
        Buffer.BltFast 885, 650, BButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        If Worker(I).Mining = 3 Then
            Buffer.BltFast 945, 650, GReturn, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
    End If
Next

For I = 1 To 5 Step 1
    If Barracks(I).Selected = True Then
        Buffer.BltFast 830, 650, AButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Buffer.BltFast 885, 650, RButton, ButtonRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
Next

End Function

Function DrawProduction()
Buffer.DrawBox 820, 750, 1020, 760
Buffer.SetFillStyle 0

For I = 1 To 20 Step 1
    If Worker(I).Selected = True Then
        Buffer.DrawBox 820, 750, 820 + Worker(I).prog, 760
    End If
Next
For I = 1 To 2 Step 1
    If TownHall(I).Selected = True Then
        Buffer.DrawBox 820, 750, 820 + TownHall(I).prog, 760
    End If
Next

Buffer.SetFillStyle 1
End Function

Function DrawBuildLoc()
For I = 1 To 20 Step 1
    With Worker(I)
        If .Construct = 1 Or .Construct = 2 Or .Construct = 3 Then
            '*************************Chk for other buildings goes here**************
            Buffer.BltFast TownHall(2).X, TownHall(2).Y, GKeep, BuildingRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
        If .Construct = 4 Or .Construct = 5 Or .Construct = 6 Then
            '*************************chk for other buildings goes here**************
            Buffer.BltFast Barracks(Barracks(1).Nextup).X, Barracks(Barracks(1).Nextup).Y, GBarracks, BuildingRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
    End With
Next

End Function

Function DrawTownHall()
For I = 1 To 2 Step 1
If TownHall(I).active = True Then
    Buffer.BltFast TownHall(I).X, TownHall(I).Y, Keep, BuildingRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
Next
End Function

Function DrawBarracks()
For I = 1 To 5 Step 1
    If Barracks(I).active = True Then
        Buffer.BltFast Barracks(I).X, Barracks(I).Y, DBarracks, BuildingRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
Next
End Function
