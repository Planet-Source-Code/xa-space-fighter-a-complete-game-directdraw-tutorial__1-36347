VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   0  'None
   Caption         =   "Space Fighter"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFrameRate 
      Interval        =   400
      Left            =   3675
      Tag             =   "Computes framerate"
      Top             =   1950
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'DirectX likes all it's variables to be predefined

'DirectDraw Related
Private dx As New DirectX7 'This is the root object. DirectDraw is created from this
Private dd As DirectDraw7 'This is DirectDraw, all things DirectDraw come from here
Private ddsd1 As DDSURFACEDESC2 'this describes the primary surface
Private primary As DirectDrawSurface7 'This surface represents the screen
Private ddsd2 As DDSURFACEDESC2 'this describes the size of the screen
Private backbuffer As DirectDrawSurface7 'Backbuffer - this is where we blt to.
Private ddsd3 As DDSURFACEDESC2 'this describes your ship
Private YourShip As DirectDrawSurface7 'Your ship surface
Private ddsd4 As DDSURFACEDESC2 'this describes enemy's ship
Private EnemyShip As DirectDrawSurface7 'Enemy ship surface
Private ddsd5 As DDSURFACEDESC2 'this describes the green laser shot
Private GreenLaser As DirectDrawSurface7 'Green laser surface
Private ddsd6 As DDSURFACEDESC2 'this describes the red laser shot
Private RedLaser As DirectDrawSurface7 'Continues in the same way...
Private ddsd7 As DDSURFACEDESC2 'stars
Private Star As DirectDrawSurface7
Private ddsd8 As DDSURFACEDESC2 'LOGO
Private HLogo As DirectDrawSurface7
Private ddFont As New StdFont '9-point green terminal

'Logic Related
Private InGame As Boolean
Private YourLocation As LocationSystem
Private Enemy() As AILogic
Private FramesFlipped As Long 'every flip should add to this
Private FrameRate As Long 'printed on screen
Private Shooting As Boolean ' are you shooting?!
Private LastBullet As Long ' when have you last shot?
Private Kills As Long ' how much have I killed
Private Health As Integer ' Am I alive?
Private Paused As Boolean 'Is the game paused?
Private Loss As Boolean ' Had I lost?
Private NextLevelScore As Long ' When do I get to the next level
Private CurLevel As Long ' Which level am I on?
Private HighScore As Long ' HighScore - All-time high
Private Score As Double ' MyScore - All-time low
Private answer As Byte ' For the story...

Private Function Apath() As String

  'Automation for getting app.path with an \ on the right

    Apath = App.Path
    If Right(Apath, 1) <> "\" Then
        Apath = Apath & "\"
    End If

End Function

Private Sub CheckLevel()

  'Levels...

    If Score >= NextLevelScore Then
        NextLevelScore = NextLevelScore * 2
        CurLevel = CurLevel + 1
        MakeShips
        Score = Score + 10
        If Health < 50 Then
            Health = Health * 2
          Else
            Health = 100
        End If
    End If

End Sub

Private Sub CheckScreen()

  Dim NeedRestore As Boolean

    Do Until ExModeActive
        DoEvents
        NeedRestore = True
    Loop
    If NeedRestore Then
        dd.RestoreAllSurfaces
        DoEvents
        InitSurfaces
    End If

End Sub

Public Function Collision(ByVal X1 As Long, ByVal Y1 As Long, ByVal W1 As Long, ByVal H1 As Long, ByVal _
                          X2 As Long, ByVal Y2 As Long, ByVal W2 As Long, ByVal H2 As Long) As Boolean

  'Checks collision between 2 objects.

    If (X1 > X2 Or X1 + W1 > X2) And X1 < X2 + W2 Then
        If (Y1 > Y2 Or Y1 + H1 > Y2) And Y1 < Y2 + H2 Then
            Collision = True
        End If
    End If

End Function

Private Sub EndIt()

  'Clearing up

    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)

    'Show Mouse NOW!
    ShowCursor (True)

    'Stop the program:
    End

End Sub

Private Sub EnemyAI(ByVal ID As Long, ByVal Level As Integer)

  'I wonder what this does...
  'The central sub for handling your enemy actions

    Select Case Level
      Case 0 To 10 '10% are fast dropping bombs
        Enemy(ID).Location.SpeedY = 10
      Case 11 To 20 '10$ are waiting for you to get near
        Enemy(ID).Shooting = False
        If Collision(Enemy(ID).Location.x, 0, ddsd4.lWidth, 2, YourLocation.x, 0, ddsd3.lWidth, 2) Then
            Enemy(ID).Shooting = True
        End If
      Case 21 To 25 ' 5% Avoidant-left
        If Enemy(ID).Location.x + ddsd4.lWidth > YourLocation.x And Enemy(ID).Location.x + ddsd4.lWidth - 4 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -1
          ElseIf Enemy(ID).Location.x + ddsd4.lWidth < YourLocation.x And Enemy(ID).Location.x + ddsd4.lWidth + 4 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 1
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
            Enemy(ID).Location.x = YourLocation.x - ddsd4.lWidth
        End If
      Case 26 To 30 ' 5% Avoidant-Right
        If Enemy(ID).Location.x - ddsd4.lWidth > YourLocation.x And Enemy(ID).Location.x - ddsd4.lWidth - 4 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -1
          ElseIf Enemy(ID).Location.x - ddsd4.lWidth < YourLocation.x And Enemy(ID).Location.x - ddsd4.lWidth + 4 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 1
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
            Enemy(ID).Location.x = YourLocation.x + ddsd4.lWidth
        End If
      Case 31 To 40 '20% are slow fighters
        Enemy(ID).Shooting = False
        If Enemy(ID).Location.x > YourLocation.x And Enemy(ID).Location.x - YourLocation.Width \ 1.4 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -1
          ElseIf Enemy(ID).Location.x < YourLocation.x And Enemy(ID).Location.x + YourLocation.Width \ 1.4 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 1
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
        End If
      Case 41 To 50 '10% always shoot
        Enemy(ID).Shooting = True
      Case 51 To 60 '20% (there will be more) are good fighters
        Enemy(ID).Shooting = False
        If Enemy(ID).Location.x > YourLocation.x And Enemy(ID).Location.x - 4 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -3
          ElseIf Enemy(ID).Location.x < YourLocation.x And Enemy(ID).Location.x + 4 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 3
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
            Enemy(ID).Location.x = YourLocation.x
        End If
      Case 61 To 70 ' 10% are ultra-fast
        Enemy(ID).Location.SpeedY = 10
        Enemy(ID).Shooting = False
        If Enemy(ID).Location.x > YourLocation.x And Enemy(ID).Location.x - YourLocation.Width \ 1.3 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -5
          ElseIf Enemy(ID).Location.x < YourLocation.x And Enemy(ID).Location.x + YourLocation.Width \ 1.3 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 5
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
        End If
      Case 71 To 80 '10% additional good fighters
        Enemy(ID).Shooting = False
        If Enemy(ID).Location.x > YourLocation.x And Enemy(ID).Location.x - 4 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -3
          ElseIf Enemy(ID).Location.x < YourLocation.x And Enemy(ID).Location.x + 4 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 3
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
            Enemy(ID).Location.x = YourLocation.x
        End If
      Case 81 To 90 '10% are floating shooters
        Enemy(ID).Location.SpeedY = 0
        Enemy(ID).Shooting = True
      Case 91 To 95 ' 5% are suicide minions
        If Enemy(ID).Location.SpeedY < 15 Then
            'sorry about the little cheating here, but it only happens once
            Enemy(ID).Location.SpeedY = 15
            Enemy(ID).Location.x = YourLocation.x - 10
            If Enemy(ID).Location.x < 0 Then
                Enemy(ID).Location.x = YourLocation.x + 10
            End If
        End If
      Case 96 To 100 ' 5% are bosses
        Enemy(ID).Location.SpeedY = 1 'bosses are slow at coming down
        'but are fast at aligning.
        If Enemy(ID).Location.x > YourLocation.x And Enemy(ID).Location.x - 5 > YourLocation.x Then
            Enemy(ID).Location.SpeedX = -5
          ElseIf Enemy(ID).Location.x < YourLocation.x And Enemy(ID).Location.x + 5 < YourLocation.x Then
            Enemy(ID).Location.SpeedX = 5
          Else
            Enemy(ID).Shooting = True
            Enemy(ID).Location.SpeedX = 0
            Enemy(ID).Location.x = YourLocation.x
        End If
    End Select

End Sub

Private Function ExModeActive() As Boolean

  'This is used to test if we're in the correct resolution.

  Dim TestCoopRes As Long

    TestCoopRes = dd.TestCooperativeLevel
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
      Else
        ExModeActive = False
    End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyY Then
        answer = 1
      ElseIf KeyCode = vbKeyN Then
        answer = 2
    End If
    If Not InGame Then
        'Only Active on the splash screen
        If KeyCode = vbKeyEscape Then
            EndIt
        End If
        If KeyCode = vbKeyReturn Then
            InGame = True
        End If
      Else
        'Only active while playing
        If KeyCode = vbKeyEscape Then
            InGame = False
            Exit Sub
        End If
        If KeyCode = vbKeyLeft Then
            YourLocation.SpeedX = -5
          ElseIf KeyCode = vbKeyRight Then
            YourLocation.SpeedX = 5
        End If
        If KeyCode = vbKeyUp Then
            YourLocation.SpeedY = -5
          ElseIf KeyCode = vbKeyDown Then
            YourLocation.SpeedY = 5
        End If
        If KeyCode = vbKeyPause Or KeyCode = vbKeyP Then
            Paused = Not Paused ' Toggle Pause
        End If
        If KeyCode = vbKeySpace Then
            Shooting = True
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        YourLocation.SpeedX = 0
      ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        YourLocation.SpeedY = 0
      ElseIf KeyCode = vbKeySpace Then
        Shooting = False
    End If

End Sub

Private Sub Form_Load()

    ShowCursor (False)

    Init

End Sub

Private Sub GameBlt()

  Dim ddrval As Long, rBack As RECT, i As Long, losefont As New StdFont

    On Local Error GoTo errOut
    CheckScreen ' to avoid some rather nasty crashes
    ddrval = backbuffer.BltColorFill(rBack, 0)
    StarFieldToBackBuffer
    Call backbuffer.DrawText(0, 0, FrameRate & "fps.", False)
    Call backbuffer.DrawText(0, 12, "Press ESC to end active game, P or Pause to toggle pause.", False)
    Call backbuffer.DrawText(0, 24, "Health left: " & Health & ". Kills: " & Kills, False)
    Call backbuffer.DrawText(0, 36, Int(Score) & " points. Current level: " & CurLevel & ". " & NextLevelScore & " required for next level", False)
    Call backbuffer.DrawText(0, 48, "High Score: " & HighScore, False)
    If Score > HighScore Then
        HighScore = Score
        SaveSetting "SpaceFighter", "High", "1", Int(Score)
    End If
    For i = 1 To UBound(Enemy)
        ddrval = backbuffer.BltFast(Enemy(i).Location.x, Enemy(i).Location.y, EnemyShip, rBack, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next i
    For i = 1 To UBound(GoodWeaponShots)
        ddrval = backbuffer.BltFast(GoodWeaponShots(i).x, GoodWeaponShots(i).y, GreenLaser, rBack, DDBLTFAST_WAIT)
    Next i
    For i = 1 To UBound(BadWeaponShots)
        ddrval = backbuffer.BltFast(BadWeaponShots(i).x, BadWeaponShots(i).y, RedLaser, rBack, DDBLTFAST_WAIT)
    Next i
    If Not Loss Then
        ddrval = backbuffer.BltFast(YourLocation.x, YourLocation.y, YourShip, rBack, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
      Else
        backbuffer.SetForeColor vbRed
        losefont.Size = 22
        losefont.Name = "Arial"
        backbuffer.SetFont losefont
        Call backbuffer.DrawText(ddsd2.lWidth / 2 - 20, ddsd2.lHeight / 2 - 10, "You lost!", False)
        backbuffer.SetFont ddFont
        Call backbuffer.DrawText(ddsd2.lWidth / 2 - 10, ddsd2.lHeight / 2 + 20, "Good luck next time!", False)
        backbuffer.SetForeColor vbGreen
    End If
    FramesFlipped = FramesFlipped + 1
    primary.Flip Nothing, DDFLIP_WAIT

errOut:
    'Just QUIT!

End Sub

Private Sub GameEngine()

    Kills = 0
    Health = 100
    CurLevel = 1
    Score = 0
    NextLevelScore = 100
    MakeShips

    Do While InGame
        If Not Paused Then
            CheckLevel
            StarFieldProcess
            MoveYourShip
            MoveAIUnits
            HandleBullets
            If Health <= 0 Then
                Loss = True
            End If
            Score = Score + 0.1
        End If
        GameBlt
        If Loss Then
            Pause (2000)
            InGame = False
            Loss = False
        End If
        DoEvents
    Loop

End Sub

Private Sub HandleBullets()

  Dim i As Long, amount As Long, cleanup As Long, mval As Long, j As Long

    On Local Error Resume Next
      amount = UBound(GoodWeaponShots)
      If Shooting And OKToShoot(LastBullet) And amount < MaxGoodBullets Then
          amount = amount + 1
          ReDim Preserve GoodWeaponShots(amount)
          GoodWeaponShots(amount).SpeedY = -5
          GoodWeaponShots(amount).x = YourLocation.x + YourLocation.Width \ 2 - 1
          GoodWeaponShots(amount).y = YourLocation.y
      End If
      For i = 1 To amount
          GoodWeaponShots(i).y = GoodWeaponShots(i).y + GoodWeaponShots(i).SpeedY
      Next i
      For i = 1 To amount
          If GoodWeaponShots(amount).y <= 0 Then
              cleanup = i
            Else
              Exit For
          End If
      Next i
      If cleanup <> 0 Then
          For j = cleanup To amount
              GoodWeaponShots(j - cleanup) = GoodWeaponShots(j)
          Next j
          ReDim Preserve GoodWeaponShots(amount - cleanup)
          cleanup = 0
      End If
      amount = UBound(BadWeaponShots)
      For i = 1 To UBound(Enemy)
          If Enemy(i).Shooting And OKToShoot(Enemy(i).LastShot) And amount < MaxBadBullets Then
              amount = amount + 1
              ReDim Preserve BadWeaponShots(amount)
              BadWeaponShots(amount).x = Enemy(i).Location.x + ddsd4.lWidth \ 2 - 1
              BadWeaponShots(amount).SpeedY = 5
              BadWeaponShots(amount).y = Enemy(i).Location.y + ddsd4.lHeight
          End If
      Next i
      For i = 1 To amount
          GoodWeaponShots(i).y = GoodWeaponShots(i).y + GoodWeaponShots(i).SpeedY
          BadWeaponShots(i).y = BadWeaponShots(i).y + BadWeaponShots(i).SpeedY
      Next i
      For i = 1 To amount
          If BadWeaponShots(amount).y >= ddsd2.lHeight Then
              cleanup = i
            Else 'NOT GOODWEAPONSHOTS(AMOUNT).Y...
              Exit For
          End If
      Next i
      If cleanup <> 0 Then
          For j = cleanup To amount
              BadWeaponShots(j - cleanup) = BadWeaponShots(j)
          Next j
          ReDim Preserve BadWeaponShots(amount - cleanup)
          cleanup = 0
      End If
      'Your shots collisions...
      For i = 1 To UBound(GoodWeaponShots)
          For j = 1 To UBound(Enemy)
              If Collision(GoodWeaponShots(i).x, GoodWeaponShots(i).y, 1, 3, Enemy(j).Location.x, Enemy(j).Location.y, ddsd4.lWidth, ddsd4.lHeight) Then
                  GoodWeaponShots(i).SpeedY = -5000
                  Enemy(j).Life = Enemy(j).Life - 1
                  Score = Score + 5
              End If
          Next j
      Next i
      'Enemy shots collisions...
      For i = 1 To UBound(BadWeaponShots)
          If Collision(BadWeaponShots(i).x, BadWeaponShots(i).y, 1, 3, YourLocation.x, YourLocation.y, YourLocation.Width, YourLocation.Height) Then
              Health = Health - 1
              BadWeaponShots(i).SpeedY = 5000
          End If
      Next i
      'SHIP COLLISIONS(!)
      For i = 1 To UBound(Enemy)
          If Collision(Enemy(i).Location.x, Enemy(i).Location.y, ddsd4.lWidth, ddsd4.lHeight, YourLocation.x, YourLocation.y, YourLocation.Width, YourLocation.Height) Then
              Health = Health - 10
              Enemy(i).Life = 0
              Score = Score - 1
          End If
      Next i
    On Local Error GoTo 0

End Sub

Private Sub Init()

  Dim caps As DDSCAPS2, i As Long

    On Local Error GoTo errOut 'If there is an error we end the program.

    Set dd = dx.DirectDrawCreate("") 'the ("") means that we want the default driver
    Me.Show 'maximises the form and makes sure it's visible

    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    'This is where we actually see a change. It states that we want a display mode
    'of 640x480 with 16 bit colour (65526 colours). the fourth argument ("0") is the
    'refresh rate. leave this to 0 and DirectX will sort out the best refresh rate. It is advised
    'that you don't mess about with this variable. the fifth variable is only used when you
    'want to use the more advanced resolutions (usually the lower, older ones)...
    Call dd.SetDisplayMode(800, 600, 16, 0, DDSDM_DEFAULT)

    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)

    'Get the backbuffer
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd2

    ' init the arrays
    ReDim GoodWeaponShots(0)
    ReDim BadWeaponShots(0)
    ' init the surfaces
    InitSurfaces

    ' init the stars
    StarFieldCreate

    ' init the font
    ddFont.Name = "Terminal"
    ddFont.Size = 9
    ddFont.Bold = False
    backbuffer.SetFont ddFont
    backbuffer.SetForeColor vbGreen

    ' init the high-score
    HighScore = GetSetting("SpaceFighter", "High", "1", 0)

    'The GAME:
    Do
        SplashEngine
        StoryEngine
        GameEngine
    Loop
    'Simple, wasn't it?

errOut:
    'If there is an error we want to close the program down straight away.
    EndIt

End Sub

Private Sub InitSurfaces()

  Dim colors As DDCOLORKEY ' for keying out colors

    'Ships

    'Those are pretty much standart
    ddsd3.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH ' use caps and define width and height
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN ' surface is not visible to user
    ddsd3.lHeight = 25
    ddsd3.lWidth = 20
    Call UniversalLoad(Apath & "Pics\yourship.bmp", ddsd3, YourShip)
    ddsd4 = ddsd3 'We can 'cheat' on this - those two descriptions are
    'identical
    Call UniversalLoad(Apath & "Pics\enemy.bmp", ddsd4, EnemyShip)
    YourLocation.Width = ddsd3.lWidth
    YourLocation.Height = ddsd3.lHeight
    'Lasers
    ddsd5.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH ' use caps and define width and height
    ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN ' surface is not visible to user
    ddsd5.lHeight = 3
    ddsd5.lWidth = 1
    Call UniversalLoad(Apath & "Pics\GreenLaser.bmp", ddsd5, GreenLaser)
    ddsd6 = ddsd5
    Call UniversalLoad(Apath & "Pics\RedLaser.bmp", ddsd6, RedLaser)
    ddsd7 = ddsd6
    ddsd7.lHeight = 1
    Call UniversalLoad(Apath & "pics\star.bmp", ddsd7, Star)
    ddsd8 = ddsd7
    ddsd8.lHeight = 100
    ddsd8.lWidth = 100
    Call UniversalLoad(Apath & "pics\logo.jpg", ddsd8, HLogo)

    colors.high = 0
    colors.low = 0
    'Black is transparent
    YourShip.SetColorKey DDCKEY_SRCBLT, colors
    EnemyShip.SetColorKey DDCKEY_SRCBLT, colors
    'We will need to blt it with SRCBLT flag for this to take effect.
    'What this does is it sets black to be not copied when blt
    'with the flag occurs.

End Sub

Private Sub MakeShips()

  Dim i As Long, x As Long

    On Error Resume Next
      x = UBound(Enemy)
    On Error GoTo 0
    ReDim Enemy(CurLevel)
    For i = x + 1 To CurLevel
        Enemy(i).Location.y = Int(ddsd2.lHeight / 3 * Rnd())
        Enemy(i).Location.SpeedY = 2
        Enemy(i).Location.x = Int(ddsd2.lWidth * Rnd())
        Enemy(i).AIID = 0 ' I like this setting - the first time the ship is around, it just falls
    Next i

End Sub

Private Sub MoveAIUnits()

  Dim i As Long, count As Long, max As Long

    For i = 1 To Rnd() * 10
        Randomize Rnd()
    Next i
    Randomize
    count = UBound(Enemy)
    For i = 1 To count
        If (Enemy(i).Location.y + ddsd4.lHeight > ddsd2.lHeight) Or Enemy(i).Life <= 0 Then
            If Enemy(i).Life <= 0 Then
                Kills = Kills + 1
            End If
            Enemy(i).Location.y = (Rnd() * ddsd3.lHeight / 3)
            Enemy(i).Location.x = Int(Rnd() * (ddsd2.lWidth - ddsd4.lWidth))
            Enemy(i).Location.SpeedY = 2 + Int(Rnd() * 3)
            Enemy(i).Location.SpeedX = 0
            max = (CurLevel * 10)
            If max > 100 Then
                max = 100
            End If
            Enemy(i).AIID = 1 + Int(Rnd() * max)
            Enemy(i).Shooting = False
            Select Case Enemy(i).AIID
              Case Is < 80: Enemy(i).Life = 1
              Case Is < 90: Enemy(i).Life = 2
              Case Is >= 90: Enemy(i).Life = 3
            End Select
        End If
        Call EnemyAI(i, Enemy(i).AIID)
        Enemy(i).Location.x = Enemy(i).Location.x + Enemy(i).Location.SpeedX
        Enemy(i).Location.y = Enemy(i).Location.y + Enemy(i).Location.SpeedY
        If Enemy(i).Location.x < 0 Then
            Enemy(i).Location.x = 0
          ElseIf Enemy(i).Location.x > ddsd2.lWidth Then
            Enemy(i).Location.x = ddsd2.lWidth - ddsd4.lWidth
        End If
        If Enemy(i).Location.y < 0 Then
            Enemy(i).Location.y = 0
          ElseIf Enemy(i).Location.y > ddsd2.lHeight Then
            Enemy(i).Location.y = ddsd2.lHeight - ddsd4.lHeight
        End If
    Next i

End Sub

Private Sub MoveYourShip()

    If YourLocation.SpeedX <> 0 Then
        YourLocation.x = YourLocation.x + YourLocation.SpeedX
        If YourLocation.x < 0 Then
            YourLocation.x = 0
          ElseIf YourLocation.x + YourLocation.Width > ddsd2.lWidth Then
            YourLocation.x = ddsd2.lWidth - YourLocation.Width
        End If
    End If
    If YourLocation.SpeedY <> 0 And AllowYMovement Then
        YourLocation.y = YourLocation.y + YourLocation.SpeedY
        If YourLocation.y < 0 Then
            YourLocation.y = 0
          ElseIf YourLocation.y + YourLocation.Height > ddsd2.lHeight Then
            YourLocation.y = ddsd2.lHeight - YourLocation.Height
        End If
    End If

End Sub

Private Function OKToShoot(LastShot As Long) As Boolean

  'Just so that we won't run into the game where you just
  'shoot aimlessly

    OKToShoot = (LastShot + 110 < GetTickCount)
    If OKToShoot Then
        LastShot = GetTickCount
    End If

End Function

Private Sub Pause(Milliseconds As Long)

  Dim old As Long

    old = GetTickCount
    Do Until old + Milliseconds < GetTickCount
        DoEvents
    Loop

End Sub

Private Sub SplashBlt()

  Dim ReturnValue As Long 'which I am going to ignore!
  Dim SrcR As RECT 'To show from where I want the picture

    CheckScreen 'Prevents some errors

    ReturnValue = backbuffer.BltColorFill(SrcR, 0)
    StarFieldToBackBuffer
    Call backbuffer.DrawText(0, 0, FrameRate & "fps.", False)
    Call backbuffer.DrawText(0, 12, "SpaceFighter v" & App.Major & "." & App.Minor, False)
    Call backbuffer.DrawText(0, 24, "Press ENTER to play, ESC to quit.", False)
    Call backbuffer.DrawText(0, 36, "High Score: " & HighScore & " points.", False)
    ReturnValue = backbuffer.BltFast(YourLocation.x, YourLocation.y, YourShip, SrcR, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    FramesFlipped = FramesFlipped + 1
    primary.Flip Nothing, DDFLIP_WAIT

End Sub

Private Sub SplashEngine()

    Do Until InGame
        StarFieldProcess
        YourLocation.y = ddsd2.lHeight - ddsd3.lHeight - 50
        YourLocation.x = (ddsd2.lWidth - ddsd3.lWidth) \ 2
        SplashBlt
        DoEvents ' Must have this
    Loop

End Sub

Private Sub StarFieldCreate()

  Dim i As Long ' index for the For ... Next loop

    Randomize
    ReDim Stars(AmountOfStars)
    For i = 1 To AmountOfStars
        Stars(i).x = Int(Rnd() * ddsd2.lWidth)
        Stars(i).y = Int(Rnd() * ddsd2.lHeight)
        Stars(i).SpeedY = 1 + Int(Rnd() * MaxStarSpeed)
    Next i

End Sub

Private Sub StarFieldProcess()

  Dim i As Long 'index for the For ... Next loop

    'Keeps our stars going!
    Randomize 'Oh, I hate those boring starfields!
    For i = 1 To AmountOfStars
        Stars(i).y = Stars(i).y + Stars(i).SpeedY
        If Stars(i).y > ddsd2.lHeight Then
            Stars(i).x = Int(Rnd() * ddsd2.lWidth)
            Stars(i).y = 0
            Stars(i).SpeedY = 1 + Int(Rnd() * MaxStarSpeed)
        End If
    Next i
    DoEvents

End Sub

Private Sub StarFieldToBackBuffer()

  Dim ReturnValue As Long, i As Long, SrcR As RECT

    For i = 1 To AmountOfStars
        ReturnValue = backbuffer.BltFast(Stars(i).x, Stars(i).y, Star, SrcR, DDBLTFAST_WAIT)
    Next i

End Sub

Private Sub StoryBlt(ByVal text As String)

  'Story Blitting...

  Dim rval As Long, r0 As RECT

    CheckScreen

    rval = backbuffer.BltColorFill(r0, 0)
    rval = backbuffer.BltFast(0, 0, HLogo, r0, DDBLTFAST_WAIT)
    Call backbuffer.DrawText(0, 100, ">" & text, False)

    primary.Flip Nothing, DDFLIP_WAIT

End Sub

Private Sub StoryEngine()

  Dim begin As Long

    'This supplies our little introduction
    Call StoryBlt("Establishing incoming connection...")
    Pause (1500)
    Call StoryBlt("Adjutant online... Reciving Transmission:")
    Pause (1200)
    Call StoryBlt(">> This is I-98 mission control to A-29 star fighter unit. Do you copy?")
    Pause (2500)
    Call StoryBlt(">> Our scouts report that a small fraction of renegade forces are gathering in sector H-38W")
    Pause (3500)
    Call StoryBlt(">> Your mission is to enter the sector and destroy the renegade forces. This is mission control I-98 out.")
    Pause (3500)
    Call StoryBlt("Transmission ended. Captain, do you need a briefing on the situation (Y/N)?")
    answer = 0
    begin = GetTickCount
    Do Until answer <> 0 Or begin + 5000 < GetTickCount
        DoEvents
    Loop
    If answer <> 2 Then
        Call StoryBlt("Proceeding with the briefing. Record J-90-387E:")
        Pause (2000)
        Call StoryBlt("> To control the ship, use your left and right arrow keys. To fire, press space.")
        Pause (3000)
        Call StoryBlt("> The ship you are commanding is much more manouvrible than any other known model.")
        Pause (3000)
        Call StoryBlt("> Enemy projectiles deal 1 unit of damage. Your ship can withstand 100.")
        Pause (3000)
        Call StoryBlt("> Avoid collisions with enemy ships - that would cost you 10 units of damage.")
        Pause (3000)
        Call StoryBlt("Record Ended. Analysis of the scout reports:")
        Pause (3000)
        Call StoryBlt("> It appears that the renegade forces are using their D-fleet. Those ships are easy to kill.")
        Pause (3000)
        Call StoryBlt("> The closer you get to the lair, the more is it guarded. Furthermore, there are different ")
        Pause (3000)
        Call StoryBlt("> models of AI controling those ships. Try to avoid getting hurt.")
        Pause (3000)
    End If
    Call StoryBlt("Transmission ended. Captain, you have full control in 3 seconds.")
    Pause (3000)

End Sub

Private Sub tmrFrameRate_Timer()

    FrameRate = FramesFlipped * 5 / 2
    FramesFlipped = 0

End Sub

Private Sub UniversalLoad(Path As String, Desc As DDSURFACEDESC2, ToSurface As DirectDrawSurface7)

  Dim loadPath As String 'The path that we will actually feed to DX
  Dim tempPic As IPictureDisp 'If we need to convert images, we use this

    'var
    Set ToSurface = Nothing 'To avoid conflicts when reloading
    If LCase$(Right$(Path, 4)) = ".bmp" Then
        loadPath = Path
      Else
        Set tempPic = LoadPicture(Path) 'Load
        SavePicture tempPic, Path & ".tmp" 'Save
        Set tempPic = Nothing ' To keep memory usage low
        loadPath = Path & ".tmp"
    End If
    Set ToSurface = dd.CreateSurfaceFromFile(loadPath, Desc) 'Actual load
    If loadPath <> Path Then
        'We created a .tmp
        Kill loadPath 'Dispose of the evidence!!!
    End If

End Sub

':) VB Code Formatter V2.12.7 (28/06/2002 11:35:27) 40 + 800 = 840 Lines
