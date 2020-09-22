VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Graphics Tutorial"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Westminster"
      Size            =   12
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7620
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   120
   End
   Begin VB.Label lblReadMe 
      Caption         =   $"frmMain.frx":030A
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image imgBBulletMask 
      Height          =   120
      Left            =   6360
      Picture         =   "frmMain.frx":0666
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgBBullet 
      Height          =   120
      Left            =   6240
      Picture         =   "frmMain.frx":0768
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgSmallRapMask 
      Height          =   375
      Left            =   5160
      Picture         =   "frmMain.frx":086A
      Top             =   4080
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgSmallRap 
      Height          =   375
      Left            =   4800
      Picture         =   "frmMain.frx":0F50
      Top             =   4080
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgExplodeMask 
      Height          =   465
      Left            =   6120
      Picture         =   "frmMain.frx":1636
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplode 
      Height          =   465
      Left            =   5640
      Picture         =   "frmMain.frx":219C
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgBullet2 
      Height          =   60
      Left            =   4320
      Picture         =   "frmMain.frx":2D02
      Top             =   2880
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgBullet1 
      Height          =   60
      Left            =   4440
      Picture         =   "frmMain.frx":2D74
      Top             =   2880
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgRapMask 
      Height          =   885
      Left            =   5760
      Picture         =   "frmMain.frx":2DE6
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRap 
      Height          =   885
      Left            =   5040
      Picture         =   "frmMain.frx":521C
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBadGuy 
      Height          =   720
      Left            =   5280
      Picture         =   "frmMain.frx":7652
      Top             =   2040
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgRapFire 
      Height          =   885
      Left            =   5760
      Picture         =   "frmMain.frx":9614
      Top             =   4680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRapFireMask 
      Height          =   885
      Left            =   5040
      Picture         =   "frmMain.frx":BA4A
      Top             =   4680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBadGuyMask 
      Height          =   720
      Left            =   4440
      Picture         =   "frmMain.frx":DE80
      Top             =   2040
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgBackGround 
      Height          =   1800
      Left            =   4680
      Picture         =   "frmMain.frx":FE42
      Top             =   120
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GRAPHICS IN VB
'By Chris George

'This project demonstrates some of Visual Basics graphics capabilities. If you really want to make a nice
'graphics program, I recommend using DirectX, but if you are just getting started this is a great way to
'learn some basics.

'Please vote or leave a comment on Planet-Source-Code if you like this program.

Option Explicit

'the constants below really affect the play of the game, because we aren't using api's or directx
'if you increase these numbers, it will dramatically affect game speed.  So try to keep these numbers small
Const MaxBadGuys = 4            'Max Number of Bad Guys on the Screen (Keep this number small)
Const MaxBullets = 25           'Max Number of Bullets you can fire
Const MaxEnemyBullets = 25      'Max Number of Bullets the enemy can fire
Const MaxAccel = 100            'Max Acceleration of Your Plane
Const EnemyFireInterval = 50    'The interval at which enemys will shoot
Const EnemySpeed = 75           'Speed at which enemies move

Public Backy As Long            'Position of the Background
Public gX As Long               'Global X position of your plane
Public gY As Long               'Global y position of your plane
Public XV As Long               'Velocity in the X direction
Public YV As Long               'Velocity in the Y direction
Public MouseX As Long           'Mouse X position
Public MouseY As Long           'Mouse Y position
Public SpaceIsPressed As Boolean            'Indicates that the space bar is pressed (fire)
Dim BadX(1 To MaxBadGuys) As Long           'Bad Guy X positions
Dim BadY(1 To MaxBadGuys) As Long           'Bad Guy Y positions
Dim BulletX(1 To MaxBullets) As Long        'BulletX positions
Dim BulletY(1 To MaxBullets) As Long        'BulletY positions
Dim BBulletX(1 To MaxEnemyBullets) As Long  'Bad Guy Bullet X positions
Dim BBulletY(1 To MaxEnemyBullets) As Long  'Bad Guy BUllet Y positions
Dim BulletCount As Long         'Number of bullets fired
Dim FireToggle As Boolean       'Toggles the flames comming out of the wings from your guns
Dim BulletToggle As Boolean     'Toggles the bullets direction
Public StartBullet As Integer   'The starting counter for bullets
Public CurrBullet As Integer    'The current bullet last fired
Public BCurrBullet As Integer   'Bad Guy current bullet last fired
Public Score As Long            'Current Score
Public Lives As Integer         'Number of Lives you have left
Public BBuletTimer As Integer   'Counter for enemy firing
Public EFireCount As Integer    'Counter for enemy firing
Public BeginGame As Boolean     'Indicates to begin the game
Public Shields As Long          'Amount of shields left on your ship

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'if the user presses the space bar then set spaceispressed to true
    Select Case KeyCode
        Case vbKeySpace
            SpaceIsPressed = True
            If BeginGame = False Then
                BeginGame = True
                gX = Me.Width / 2 - Me.imgRap.Width / 2
                gY = Me.Height - Me.imgRap.Height - 300
                For i = 1 To MaxBadGuys         'position of the bad guys
                    BadX(i) = Rnd * Me.Width
                    BadY(i) = (Rnd * Me.Height) - Me.Height
                Next
                StartBullet = 1
                Lives = 3
                For i = 1 To MaxBullets
                    BulletY(i) = -100
                Next
                For i = 1 To MaxEnemyBullets
                    BBulletY(i) = Me.Height
                Next
                SpaceIsPressed = False
                Me.Shields = 4000
                Score = 0
            End If
        Case vbKeyF2
            BeginGame = True
            gX = Me.Width / 2 - Me.imgRap.Width / 2
            gY = Me.Height - Me.imgRap.Height - 300
            For i = 1 To MaxBadGuys         'position of the bad guys
                BadX(i) = Rnd * Me.Width
                BadY(i) = (Rnd * Me.Height) - Me.Height
            Next
            StartBullet = 1
            Lives = 3
            For i = 1 To MaxBullets
                BulletY(i) = -100
            Next
            For i = 1 To MaxEnemyBullets
                BBulletY(i) = Me.Height
            Next
            SpaceIsPressed = False
            Me.Shields = 4000
            Me.Timer1.Interval = 1
            Score = 0
        Case vbKeyP
            If Me.Timer1.Interval = 0 Then
                Me.Timer1.Interval = 1
            Else
                Me.Timer1.Interval = 0
                Me.FontSize = 28
                Me.CurrentX = Me.Width / 2 - Me.TextWidth("Paused") / 2
                Me.CurrentY = Me.Height / 2 - 2000
                Me.Print "Paused"
            End If
            
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'if the user releases the space bar then set spaceispressed to false
    Select Case KeyCode
        Case vbKeySpace
            SpaceIsPressed = False
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'initalize variables
    StartBullet = 1                 'start counter for bullets
    Lives = 3                       'starting number of lives
    gX = Me.imgRapFire.Left         'position of your plane
    gY = Me.imgRapFire.Top
    For i = 1 To MaxBadGuys         'position of the bad guys
        BadX(i) = Rnd * Me.Width
        BadY(i) = Rnd * Me.Height
    Next
    Me.DrawStyle = 0
    Me.FillStyle = 0
    Me.FillColor = vbRed
    Me.Shields = 4000
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if the user presses the mouse button, it does the same thing a pressing space so
    'set spaceispressed to true to start firing
    SpaceIsPressed = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'get the current mouse position
    MouseX = X
    MouseY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if the user releases the mouse button the set spaceispressed to false to stop firing
    SpaceIsPressed = False
End Sub

Private Sub Timer1_Timer()
    '==========MAIN GAME LOOP=============
    
    Dim i As Integer    'counter
    Dim j As Integer    'counter
    Dim TempX As Long
    Dim TempY As Long
    
    'First clear the screen
    Me.Cls
    
    'Check to see if your ship is at the mouse position
    'if not then increase its velocity to go to that position
    'this creats a gliding effect instead of a dead halt
    If MouseX > gX Then XV = XV + 10
    If MouseX < gX Then XV = XV - 10
    If MouseY > gY Then YV = YV + 10
    If MouseY < gY Then YV = YV - 10

    'check to make sure your ship isn't accelerating out of control
    If XV > MaxAccel Then XV = MaxAccel
    If -XV > MaxAccel Then XV = -MaxAccel
    If YV > MaxAccel Then YV = MaxAccel
    If -YV > MaxAccel Then YV = -MaxAccel
    
    'move your ship at its current velocity
    gX = gX + XV
    gY = gY + YV
    
    'if your ship goes to the edge of the screen, stop its movement and set its velocity to 0
    If gX < 0 Then gX = 0: XV = 0
    If gY < 0 Then gY = 0: YV = 0
    If gX > Me.Width - Me.imgRap.Width Then gX = Me.Width - Me.imgRap.Width: XV = 0
    If gY > Me.Height - Me.imgRap.Height Then gY = Me.Height - Me.imgRap.Height: YV = 0
    
    'move the background
    Backy = Backy + 25
    If Backy > Me.imgBackGround.Height Then Backy = 0
    'tile the background accross the form
    For i = 1 To Me.Width Step Me.imgBackGround.Width
        For j = -(Me.imgBackGround.Height) To Me.Height Step Me.imgBackGround.Height
            Me.PaintPicture Me.imgBackGround.Picture, i, j + Backy
        Next
    Next
    
    'if the user pressed the space bar then start firing
    If SpaceIsPressed And FireToggle = True Then
        'increment the current bullet
        CurrBullet = CurrBullet + 1
        'make sure the current bullet isn't greater then the max number of bullets
        If CurrBullet > MaxBullets Then CurrBullet = 1: StartBullet = 1
        'set the bullets position to your ships position
        BulletX(CurrBullet) = gX
        BulletY(CurrBullet) = gY
        'paint the mask picture to your ships position
        Me.PaintPicture Me.imgRapFireMask.Picture, gX, gY, , , , , , , vbSrcAnd
        'paint the picture to your ships position, this creates the transparent effect
        Me.PaintPicture Me.imgRapFire.Picture, gX, gY, , , , , , , vbSrcPaint
        'toggle the firing to false so it changes the appearence of the flames on the side of your ship from firing
        FireToggle = False
    Else
        'paint the mask picture to your ships position
        Me.PaintPicture Me.imgRapMask.Picture, gX, gY, , , , , , , vbSrcAnd
        'paint the picture to your ships position, this creates the transparent effect
        Me.PaintPicture Me.imgRap.Picture, gX, gY, , , , , , , vbSrcPaint
        'toggle the firing to true so it changes the appearence of the flames on the side of your ship from firing
        FireToggle = True
    End If
    
    'loop through your bullets and check to see if they hit any bad guys
    For i = StartBullet To MaxBullets
        'move the bullets
        BulletY(i) = BulletY(i) - 125
        'check to make sure the bullet is still on the screen, if not don't draw it to save time
        If BulletY(i) > 0 Then
            If BulletToggle = True Then
                'draw the left bullet mask
                Me.PaintPicture Me.imgBullet1.Picture, BulletX(i) + Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcAnd
                'draw the left bullet
                Me.PaintPicture Me.imgBullet2.Picture, BulletX(i) + Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcPaint
                'draw the right bullet mask
                Me.PaintPicture Me.imgBullet1.Picture, BulletX(i) + Me.imgRap.Width - Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcAnd
                'draw the right bullet
                Me.PaintPicture Me.imgBullet2.Picture, BulletX(i) + Me.imgRap.Width - Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcPaint
                'toggle bullettoggle so it changes the appearence of the bullets
                BulletToggle = False
            Else
                Me.PaintPicture Me.imgBullet2.Picture, BulletX(i) + Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcAnd
                Me.PaintPicture Me.imgBullet1.Picture, BulletX(i) + Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcPaint
                Me.PaintPicture Me.imgBullet2.Picture, BulletX(i) + Me.imgRap.Width - Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcAnd
                Me.PaintPicture Me.imgBullet1.Picture, BulletX(i) + Me.imgRap.Width - Me.imgBullet1.Width, BulletY(i), , , , , , , vbSrcPaint
    
                BulletToggle = True
            End If
            
            'loop through the bad guys and check to see if any bullets hit them
            For j = 1 To MaxBadGuys
                If BulletY(i) > BadY(j) And BulletY(i) < (BadY(j) + Me.imgBadGuy.Height) And BulletY(i) > 0 Then
                    If BulletX(i) > BadX(j) - Me.imgRap.Width And BulletX(i) < (BadX(j) + Me.imgBadGuy.Width) Then
                        'if a bullet hit then draw the explosion mask
                        Me.PaintPicture Me.imgExplodeMask.Picture, BadX(j), BadY(j), , , , , , , vbSrcAnd
                        'then draw the explosion picture
                        Me.PaintPicture Me.imgExplode.Picture, BadX(j), BadY(j), , , , , , , vbSrcPaint
                        'set the bad guys position back to the top
                        BadY(j) = -(Me.imgBadGuy.Height)
                        BadX(j) = Rnd * Me.Width
                        'set the bullets position to beyond the top so it doesn't get drawn
                        BulletY(i) = -100
                        'add 100 to the score
                        If BeginGame = True Then Score = Score + 100
                    End If
                End If
            Next
        End If
    Next

    
    'loop through the bad guys and fire a bullet if its there turn
    For i = 1 To MaxBadGuys
        'increment the enemy fire count
        EFireCount = EFireCount + 1
        'move the bad guys down
        BadY(i) = BadY(i) + EnemySpeed
        'comment out the next two lines if you don't want the bad guys following you
        If BadX(i) > gX Then BadX(i) = BadX(i) - 25
        If BadX(i) < gX Then BadX(i) = BadX(i) + 25
        
        'if the bad guy goes off the screen then move it to the top
        If BadY(i) > Me.Height Then BadY(i) = -Me.imgBadGuy.Height: BadX(i) = Rnd * Me.Width
        'check to make sure the bad guy is on the screen, if not don't draw it
        If BadY(i) < Me.Height Then
            'paint the mask of the badguy
            Me.PaintPicture Me.imgBadGuyMask, BadX(i), BadY(i), , , , , , , vbSrcAnd
            'paint the badguy
            Me.PaintPicture Me.imgBadGuy, BadX(i), BadY(i), , , , , , , vbSrcPaint
        End If
        'if its the enemys turn to fire (a random interval over the enemy fire interval)
        If EFireCount > (EnemyFireInterval + Rnd * 100) Then
            'increment the enemy bullet count
            BCurrBullet = BCurrBullet + 1: EFireCount = 1
            'make sure the bullet count isnt over the max
            If BCurrBullet > MaxEnemyBullets Then BCurrBullet = 1
            'set the enemy bullet to the middle of the bad guy
            BBulletX(BCurrBullet) = BadX(i) + Me.imgBadGuy.Width / 2 - Me.imgBBullet.Width / 2
            BBulletY(BCurrBullet) = BadY(i)
        End If
        'check to see if you ran into the bad guy
        If BeginGame = True Then
            If BadY(i) + Me.imgBadGuy.Height > gY And BadY(i) < gY + Me.imgRap.Height Then
                If BadX(i) + Me.imgBadGuy.Width > gX And BadX(i) < gX + Me.imgRap.Width Then
                    Shields = Shields - 400
                    For j = 1 To 100
                        TempX = BadX(i) + Rnd * Me.imgBadGuy.Width
                        TempY = BadY(i) + Rnd * Me.imgBadGuy.Height
                        Me.PaintPicture Me.imgExplodeMask, TempX, TempY, , , , , , , vbSrcAnd
                        Me.PaintPicture Me.imgExplode, TempX, TempY, , , , , , , vbSrcPaint
                    Next
                    'if shields get less than 100 then blow up and subtract 1 life
                    If Shields <= 0 Then
                        Lives = Lives - 1
                        For j = 1 To 250
                            TempX = gX + Rnd * Me.imgRap.Width
                            TempY = gY + Rnd * Me.imgRap.Height
                            Me.PaintPicture Me.imgExplodeMask, TempX, TempY, , , , , , , vbSrcAnd
                            Me.PaintPicture Me.imgExplode, TempX, TempY, , , , , , , vbSrcPaint
                        Next
                        'if lives is less than 0 then game over
                        If Lives = -1 Then
                            Me.Timer1.Interval = 0
                            Me.FontSize = 32
                            Me.CurrentX = Me.Width / 2 - Me.TextWidth("GAME OVER") / 2
                            Me.CurrentY = Me.Height / 2 - 2000
                            Me.Print "GAME OVER"
                            Me.FontSize = 28
                            Me.Print ""
                            Me.CurrentX = Me.Width / 2 - Me.TextWidth("Press F2 To Start Over") / 2
                            Me.Print "Press F2 To Start Over"
                        End If
                        
                        Shields = 4000
                    End If
                    BadY(i) = Me.Height
                End If
            End If
        End If
    Next
    
    'loop through the enemy bullets and move them
    For i = 1 To MaxEnemyBullets
        'move the enemy bullet
        BBulletY(i) = BBulletY(i) + EnemySpeed + 25
        'comment these 2 lines out to make the bullets not go toward your ship
        If BBulletX(i) > gX Then BBulletX(i) = BBulletX(i) - 25
        If BBulletX(i) < gX Then BBulletX(i) = BBulletX(i) + 25
        'make sure its still on the screen, if not don't draw it
        If BBulletY(i) < Me.Height Then
            If BeginGame = True Then
                If BBulletY(i) < gY + Me.imgRap.Height And BBulletY(i) > gY Then
                    If BBulletX(i) < gX + Me.imgRap.Width And BBulletX(i) > gX Then
                        Shields = Shields - 100
                        'if shields get less than 100 then blow up and subtract 1 life
                        If Shields <= 0 Then
                            Lives = Lives - 1
                            For j = 1 To 250
                                TempX = gX + Rnd * Me.imgRap.Width
                                TempY = gY + Rnd * Me.imgRap.Height
                                Me.PaintPicture Me.imgExplodeMask, TempX, TempY, , , , , , , vbSrcAnd
                                Me.PaintPicture Me.imgExplode, TempX, TempY, , , , , , , vbSrcPaint
                            Next
                            'if lives is less than 0 then game over
                            If Lives = -1 Then
                                Me.Timer1.Interval = 0
                                Me.FontSize = 32
                                Me.CurrentX = Me.Width / 2 - Me.TextWidth("GAME OVER") / 2
                                Me.CurrentY = Me.Height / 2 - 2000
                                Me.Print "GAME OVER"
                                Me.FontSize = 28
                                Me.Print ""
                                Me.CurrentX = Me.Width / 2 - Me.TextWidth("Press F2 To Start Over") / 2
                                Me.Print "Press F2 To Start Over"
                            End If
                            
                            Shields = 4000
                        End If
                        Me.PaintPicture Me.imgExplodeMask, BBulletX(i), BBulletY(i), , , , , , , vbSrcAnd
                        Me.PaintPicture Me.imgExplode, BBulletX(i), BBulletY(i), , , , , , , vbSrcPaint
                        BBulletY(i) = Me.Height
                    End If
                End If
            End If
            Me.PaintPicture Me.imgBBulletMask, BBulletX(i), BBulletY(i), , , , , , , vbSrcAnd
            Me.PaintPicture Me.imgBBullet, BBulletX(i), BBulletY(i), , , , , , , vbSrcPaint
        End If
    Next
    
    'paint the number of lives under the score
    For i = 1 To Lives
        Me.PaintPicture Me.imgSmallRapMask.Picture, 100 + (i * (Me.imgSmallRap.Width + 25)), 400, , , , , , , vbSrcAnd
        Me.PaintPicture Me.imgSmallRap.Picture, 100 + (i * (Me.imgSmallRap.Width + 25)), 400, , , , , , , vbSrcPaint
    Next
    
    'if the game hasn't started yet print instructions up on the screen
    If BeginGame = False Then
        Me.FontSize = 16
        Me.CurrentX = Me.Width / 2 - Me.TextWidth("Raptor Demo in VB") / 2
        Me.CurrentY = Me.Height / 2 - 2000
        Me.Print "Raptor Demo in VB"
        Me.CurrentX = Me.Width / 2 - Me.TextWidth("Press Space to Begin") / 2
        Me.Print "Press Space to Begin"
        Me.CurrentX = Me.Width / 2 - Me.TextWidth("Use Your Mouse To Control Your Ship") / 2
        Me.Print "Use Your Mouse To Control Your Ship"
        Me.CurrentX = Me.Width / 2 - Me.TextWidth("Press The Mouse Button To Fire") / 2
        Me.Print "Press The Mouse Button To Fire"
    End If
    
    'paint your sheilds at the bottom
    Me.Line (100, (Me.Height - 1000) - Shields)-(300, Me.Height - 1000), vbRed, B
    
    Me.FontSize = 12
    'put the score at the top and my name at the bottom of the screen
    Me.CurrentX = 100
    Me.CurrentY = 100
    Me.Print "Score: " & Score
    Me.CurrentX = Me.Width - 2500
    Me.CurrentY = Me.Height - 800
    Me.Print "Created By: Chris George"
    
    'if the user has paused then show the word in the center of the screen
    If Me.Timer1.Interval = 0 And Lives >= 0 Then
        Me.FontSize = 28
        Me.CurrentX = Me.Width / 2 - Me.TextWidth("Paused") / 2
        Me.CurrentY = Me.Height / 2 - 2000
        Me.Print "Paused"
    End If
    
End Sub
