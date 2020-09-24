VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2130
      Top             =   1785
   End
   Begin VB.Timer tmrPinger 
      Interval        =   75
      Left            =   1095
      Top             =   1665
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'DirectX likes all it's variables to be predefined
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Type Location
    x As Long
    y As Long
    SpeedX As Single
    SpeedY As Single
End Type
Dim binit As Boolean 'A simple flag (true/false) that states whether we've initialised or not. If the initialisation is successful 'this changes to true, the program also checks before doing any drawing if this flag is true. If the initialisation failed and we 'try and draw things we'll get lots of errors...
Private XBackColor As Long
Dim dx As New DirectX7 'This is the root object. DirectDraw is created from this
Dim dd As DirectDraw7 'This is DirectDraw, all things DirectDraw come from here
Dim Mainsurf As DirectDrawSurface7 'This holds our bitmap
Dim primary As DirectDrawSurface7 'This surface represents the screen - see earlier in the tutorial
Dim backbuffer As DirectDrawSurface7 'This was mentioned earlier on...
Dim Pinger As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2 'this describes the primary surface
Dim ddsd2 As DDSURFACEDESC2 'this describes the ball
Dim ddsd3 As DDSURFACEDESC2 'this describes the size of the screen
Dim ddsd4 As DDSURFACEDESC2 'this describes the pinger
Private TLast As Long
Private bltCount As Long
Private score As Long
Private HighScore As Long
Dim framerate As Integer
Private Ball As Location
Private PingerLoc As Location
Dim Loss As Boolean ' flag to show that the ball was lost
Dim ddFont As New StdFont
Dim ddLoseFont As New StdFont
Dim brunning As Boolean 'this is another flag that states whether or not the main game loop is running.
Dim CurModeActiveStatus As Boolean 'This checks that we still have the correct display mode
Dim bRestore As Boolean 'If we don't have the correct display mode then this flag states that we need to restore the display mode
Private InGame As Boolean
Private CurLevel As Long
Private Lives As Long
Private NextLevel As Long
Private PauseX As Boolean

Sub Init()

  Dim count As Long, S As Integer, res As String, h As Long, w As Long, d As Long

    On Local Error GoTo errOut 'If there is an error we end the program.

    CurLevel = 1
    Set dd = dx.DirectDrawCreate("") 'the ("") means that we want the default driver
    Me.Show 'maximises the form and makes sure it's visible

    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    res = GetSetting("Ping", "Res", "Res", "800x600x16")
    w = XTract(res, 1)
    h = XTract(res, 2)
    d = XTract(res, 3)

    Call dd.SetDisplayMode(w, h, d, 0, DDSDM_DEFAULT)
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)

    'Get the backbuffer
  Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd3

    ' init the surfaces
    InitSurfaces

    'Do the fonts
    ddFont.Name = "Terminal"
    ddFont.Size = 9
    backbuffer.SetForeColor RGB(55, 255, 55)
    backbuffer.SetFontTransparency True
    backbuffer.SetFont ddFont

    HighScore = GetSetting("Balls", "High", "1", 0)
    'Main Program Loop - show splash, and play game... REPEAT
    Do
        SplashEngine
        GameEngine
    Loop
errOut:
    'If there is an error we want to close the program down straight away.
    endit

End Sub

Public Sub SplashEngine()

    InGame = False
    Randomize
    Do Until InGame
        PingerLoc.x = (Rnd() * (ddsd3.lWidth - ddsd4.lWidth) \ 1)
        PingerLoc.y = (Rnd() * (ddsd3.lHeight - ddsd4.lHeight) \ 1)
        Ball.x = (Rnd() * (ddsd3.lWidth - ddsd2.lWidth) \ 1)
        Ball.y = (Rnd() * (ddsd3.lHeight - ddsd2.lHeight) \ 1)
        SplashBlt
        Pause (500)
        DoEvents
    Loop

End Sub

Public Sub SplashBlt()

  Dim rect2 As RECT

    CheckScreen
    backbuffer.BltColorFill rect2, 0
    backbuffer.DrawText 0, 0, "Press ENTER to play", False
    backbuffer.DrawText 0, 10, "High Score: " & HighScore, False
    backbuffer.BltFast PingerLoc.x, PingerLoc.y, Pinger, rect2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    backbuffer.BltFast Ball.x, Ball.y, Mainsurf, rect2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    primary.Flip Nothing, DDFLIP_WAIT

End Sub

Public Function XTract(ByVal S As String, ByVal ID As Long) As String

  Dim x As Long, i As Long, xs As Long

    S = S & "x"
    xs = 1
    x = InStr(xs + 1, S, "x")
    For i = 1 To ID - 1
        xs = x
        x = InStr(xs + 1, S, "x")
    Next i
    XTract = Replace(Mid(S, xs, x - xs), "x", "")

End Function

Sub InitSurfaces()

    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH 'default flags
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = 30
    ddsd2.lHeight = 30
    Call UniversalLoad(App.Path & "\ball" & CurLevel & ".jpg", Mainsurf, ddsd2)
    ddsd4.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd4.lWidth = 73
    ddsd4.lHeight = 21
    Call UniversalLoad(App.Path & "\pinger" & CurLevel & ".jpg", Pinger, ddsd4)

  Dim key As DDCOLORKEY
    key.low = 0
    key.high = 0
    Pinger.SetColorKey DDCKEY_SRCBLT, key
    Mainsurf.SetColorKey DDCKEY_SRCBLT, key

End Sub

Public Sub GameEngine()

  Dim needBall As Boolean, d As Long

    CurLevel = 1
    NextLevel = 2000
    InitSurfaces
    Lives = 3
    binit = True
    brunning = True
    Do Until ExModeActive
        DoEvents
    Loop
    Ball.SpeedX = 3
    Ball.SpeedY = 3
    Ball.y = 0
    Ball.x = 0
    PingerLoc.y = ddsd3.lHeight - ddsd4.lHeight - 30
    PingerLoc.x = (ddsd3.lWidth - ddsd4.lWidth) / 2
    Do While InGame
        MovePinger
        Ball.x = Ball.x + Ball.SpeedX
        Ball.y = Ball.y + Ball.SpeedY
        If Ball.x + ddsd2.lWidth >= ddsd3.lWidth Or Ball.x <= 0 Then
            Ball.SpeedX = -Ball.SpeedX
            score = score + 15
        End If
        If Ball.y <= 0 Then IncreaseBallSpeed (True)
        If Ball.y + ddsd2.lHeight >= ddsd3.lHeight Then
            Lives = Lives - 1
            needBall = True
            If Lives = 0 Then Loss = True
        End If
        If Ball.y + ddsd2.lHeight > PingerLoc.y And _
           Ball.y < PingerLoc.y + ddsd4.lHeight And _
           Ball.x > PingerLoc.x And _
           Ball.x + ddsd2.lWidth < PingerLoc.x + ddsd4.lWidth Then
            If Ball.SpeedY > 0 Then
                IncreaseBallSpeed (True)
                Ball.y = Ball.y + 2 * Ball.SpeedY
            End If
        End If
        GameBlt
        If Loss Then
            d = Second(Now) + 2
            If d >= 60 Then d = d - 60
            Beep
            Pause (1000)
            Loss = False
            InGame = False
            Exit Sub
        End If
        If needBall Then
            Ball.y = 0
            Ball.x = 0
            PingerLoc.x = (ddsd3.lWidth - ddsd4.lWidth) / 2
            needBall = False
            Ball.SpeedX = 3
            Ball.SpeedY = 3
            Pause (1500)
        End If
        If PauseX Then
            Do While PauseX
                DoEvents
            Loop
        End If
        DoEvents 'you MUST have a doevents in the loop, otherwise you'll overflow the
        'system (which is bad). All your application does is keep sending messages to DirectX
        'and windows, if you dont give them time to complete the operation they'll crash.
        'adding doevents allows windows to finish doing things that its doing.
    Loop

End Sub

Private Sub CheckScreen()

    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces
    End If

End Sub

Public Sub GameBlt()

  Dim timeDiff As Long
  Dim rBack As RECT, rFront As RECT
  Dim ddrval As Long

    On Local Error GoTo errOut
    If binit = False Then Exit Sub

    ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    CheckScreen

    rBack.Bottom = 0
    rBack.Right = 0
    ddrval = backbuffer.BltColorFill(rBack, XBackColor)
    Call backbuffer.DrawText(0, 0, framerate & "fps. Score: " & score & ". High: " & HighScore & ". " & Lives-1 & " Ball(s) Left", False)
    Call backbuffer.DrawText(0, 10, "Current Level: " & CurLevel & ". Next Level: " & NextLevel & " points.", False)
    Call backbuffer.DrawText(0, 20, "Press ESC to end game.", False)
    Call backbuffer.DrawText(0, 31, "Press 'P' or 'Pause' to pause.", False)

    ddrval = backbuffer.BltFast(PingerLoc.x, PingerLoc.y, Pinger, rBack, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    If Not Loss Then
        ddrval = backbuffer.BltFast(Ball.x, Ball.y, Mainsurf, rBack, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End If
    If score > HighScore Then
        HighScore = score
        SaveSetting "Balls", "High", "1", HighScore
    End If

    If Loss Then
        backbuffer.SetForeColor vbRed
        backbuffer.DrawText ddsd3.lWidth / 2 - 50, ddsd3.lHeight / 2 - 10, "You lost!", False
        backbuffer.SetForeColor vbGreen
        score = 0
    End If
    bltCount = bltCount + 1

    'flip the back buffer to the screen
    primary.Flip Nothing, DDFLIP_WAIT
    'At this point we have completed one cycle, and we can now see something on screen

errOut:

End Sub

Sub endit()

  'This procedure is called at the end of the loop, or whenever there is an error.
  'Although you can get away without these few lines it is a good idea to keep them
  'as you can get unpredictable results if you leave windows to "clear-up" after you.

  'This line restores you back to your default (windows) resolution.

    Call dd.RestoreDisplayMode
    'This tells windows/directX that we no longer want exclusive access
    'to the graphics features/directdraw
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    'Stop the program:
    End

End Sub

Private Sub Form_Click()

  'Clicking the form will result in the program closing down.
  'because the form is maximised (and therefore covers the whole screen)
  'where you click is not important.

    If Not InGame Then
        endit
    End If
    
End Sub

Private Sub Form_Initialize()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If InGame Then
    If KeyCode = vbKeyRight Then
        If PingerLoc.SpeedX <= 0 Then
            PingerLoc.SpeedX = 3 + CurLevel * 2
        End If
      ElseIf KeyCode = vbKeyLeft Then
        If PingerLoc.SpeedX >= 0 Then
            PingerLoc.SpeedX = -3 - CurLevel * 2
        End If
      ElseIf KeyCode = vbKeyUp And CurLevel > 4 Then
            PingerLoc.SpeedY = -CurLevel
      ElseIf KeyCode = vbKeyDown And CurLevel > 4 Then
            PingerLoc.SpeedY = CurLevel
    End If     
    End If
    If KeyCode = vbKeyReturn Then
        InGame = True
    End If
    If KeyCode = vbKeyP Or KeyCode = vbKeyPause Then
        PauseMode
    End If
    If KeyCode = vbKeyEscape And Not PauseX And InGame Then
        InGame = False
      ElseIf KeyCode = vbKeyEscape And Not InGame Then
        endit
    End If

End Sub

Private Sub PauseBlt()

    CheckScreen

    backbuffer.DrawText 0, 41, "Paused. Press 'P' or 'Pause' to play", False
    primary.Flip Nothing, DDFLIP_WAIT

End Sub

Private Sub PauseMode()

    PauseX = Not PauseX
    If PauseX Then
        backbuffer.SetFontTransparency False
        Do While PauseX
            PauseBlt
            DoEvents
        Loop
      Else
        backbuffer.SetFontTransparency True
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    PingerLoc.SpeedX = 0
    PingerLoc.SpeedY = 0

End Sub

Private Sub Form_Load()

  'Starts the whole program.

    Init

End Sub

Private Sub Form_Paint()

  'If windows sends a "paint" message translate this into a call
  'to DirectDraw.

End Sub

Function ExModeActive() As Boolean

  'This is used to test if we're in the correct resolution.

  Dim TestCoopRes As Long

    TestCoopRes = dd.TestCooperativeLevel
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
      Else
        ExModeActive = False
    End If

End Function

Public Sub UniversalLoad(Picture As String, Surface As DirectDrawSurface7, Description As DDSURFACEDESC2)

  Dim x As IPictureDisp, Path As String

    On Local Error GoTo endit
    Set Surface = Nothing
    If StrComp(Right(Picture, 4), ".bmp", vbTextCompare) = 0 Then
        Path = Picture
      Else
        Set x = LoadPicture(Picture)
        Call SavePicture(x, Picture & ".tmp")
        Set x = Nothing
        Path = Picture & ".tmp"
    End If
    Set Surface = dd.CreateSurfaceFromFile(Path, Description)
    If StrComp(Right(Picture, 4), ".bmp", vbTextCompare) = 1 Then
        Kill Path
    End If

Exit Sub

endit:
    endit

End Sub

Public Sub MovePinger()

    If PingerLoc.SpeedX <> 0 Then
        If PingerLoc.x + PingerLoc.SpeedX >= 0 And _
           PingerLoc.x + ddsd4.lWidth + PingerLoc.SpeedX <= ddsd3.lWidth Then
            PingerLoc.x = PingerLoc.x + PingerLoc.SpeedX
        End If
    End If

End Sub

Private Sub Timer1_Timer()

    If Not PauseX Then
        framerate = bltCount * 5
        bltCount = 0
    End If

End Sub

Private Sub tmrPinger_Timer()

    If Not PauseX Then
        If Ball.SpeedY < 0 Then
        score = score - Ball.SpeedY
        Else
        score = score + Ball.SpeedY
        End If
        If score > NextLevel And CurLevel < 7 Then
            NextLevel = NextLevel * 2
            CurLevel = CurLevel + 1
            InitSurfaces
            Lives = Lives + 1
        End If
    End If

End Sub

Private Sub Pause(Milli As Long)

  Dim t As Long

    t = GetTickCount
    Do Until t + Milli <= GetTickCount
        DoEvents
    Loop

End Sub

Private Sub IncreaseBallSpeed(ByVal Invert As Boolean)

    Ball.SpeedY = Ball.SpeedY * IIf(Invert, -1.015, 1.015)
    Ball.SpeedX = Ball.SpeedX * 1.016
    score = score + 30

End Sub

':) VB Code Formatter V2.12.7 (24/06/2002 13:51:45) 38 + 445 = 483 Lines
