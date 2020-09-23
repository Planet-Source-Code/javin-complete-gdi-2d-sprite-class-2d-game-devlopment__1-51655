VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Used to just grab framerates.
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private sprSleep As Sprite2D
Private sprBoard As Sprite2D
Private sprPiece As Sprite2D
Private sprSquare As Sprite2D
Private sprBuffer As Sprite2D
Private sprSprite As Sprite2D
Private sprRunner As Sprite2D

Private booIsRunning As Boolean
Private booUseSleep As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If "Escape" is pressed (Chr(27)) then quit running the loop.
    If KeyCode = 27 Then booIsRunning = False
    'If we've already quit the loop, ignore everything following.
    If Not booIsRunning Then Exit Sub
    'Toggles the "sleep" mode on, for testing the load on the CPU.
    If KeyCode = 83 Then
        If booUseSleep Then
            booUseSleep = False
            sprSleep.EraseSprite
        Else
            booUseSleep = True
            sprSleep.DrawSprite
        End If
    End If
End Sub

Private Sub Form_Activate()
    'Always best to run your game loop in the Activate event.  This way the form has loaded
    'and is ready to take over the .Paint event for you.  Manually calling the .Refresh for
    'the form is painfully slow.
    RunGameLoop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Even something so simple as a mouse-click will shut this bad boy down!
    booIsRunning = False
End Sub
 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not booIsRunning Then Exit Sub
    'If the mouse is moved, set the middle of the sprite's location to the mouse's position.
    sprSprite.SetPosition (x / Screen.TwipsPerPixelX) - (sprSprite.Cell_Width / 2), (y / Screen.TwipsPerPixelY) - (sprSprite.Cell_Height / 2)
End Sub

Private Sub RunGameLoop()
    'Create all of your objects this way.  Creating them in the global declarations with:
    'Private Blah As New Sprite 2D
    'is a MASSIVE slowdown, since it actually destroys, and recreates the sprite whenever
    'it loses focus.  (In this particular example, that wouldn't be much of an issue, but
    'trust me, it's a bad habit to get into.)
    Set sprBoard = New Sprite2D
    Set sprSquare = New Sprite2D
    Set sprBuffer = New Sprite2D
    Set sprPiece = New Sprite2D
    Set sprSprite = New Sprite2D
    Set sprSleep = New Sprite2D
    Set sprRunner = New Sprite2D
    
    Dim lngCount As Long
    Dim lngFPS As Long, lngTick As Long
    
    'For maximum exposure!  (.)(.)
    Me.WindowState = vbMaximized
    
    booIsRunning = True
    
    'Hide the mouse.  This is actually the CORRECT way to use the ShowCursor API, since
    'the number of times you show it is incremental.  Thus, if I hide it 5 times, I have
    'to show it 5 times in order to see it, and you have NO idea what other apps have done
    'with it.  So use this method to show/hide your cursor.
    'While ShowCursor(0) > 0: Wend
    
    'Load all of the images
    'The .Parent_hDC property is where the sprite will be drawn to by default.
    'We set the buffer so that it's the only sprite that draws to the screen (or
    'the form in this case) by default.  Make sure to always load a sprite before
    'trying to tell another sprite to draw to it.  Additionally, using a
    'backbuffer also means you don't have to dick with the Form.Refresh function and
    'autoredraw properties.  This will in some cases more than double your FPS.
    
    'First, we load the BackBuffer, telling it to draw to the form.  Note that this is
    'the ONLY sprite that is specifically told NOT to use GDI.  For some reason,
    'sprites that have to be BitBltted directly to the screen have a HUGE slowdown
    'if they use GDI.  GDI is faster, however with "hidden" DC->DC BitBlts.
    'If ya don't believe me, set this part to "True" and see what happens.
    sprBuffer.LoadEmptySprite Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, Me.hdc, , , False, RGB(128, 0, 128)
    
    'Then, we'll load the board and set it's "parent" to the buffer.  For images like
    'this, GDI is actually faster than bitmaps created using "CreateCompatibleBitmap."
    'I didn't bother including a "UseGDI" flag here, since there should never be an
    'excuse to draw an "image" directly to the screen.  A buffer should ALWAYS be used
    '(like the one above) to keep "flicker" down.
    sprBoard.LoadImageSprite App.Path & "\Board.jpg", sprBuffer.hdc
        
    'This will be used to show when we do and don't use "sleep."
    sprSleep.LoadBMPSprite App.Path & "\Sleep.bmp", sprBuffer.hdc, , , 0
    
    'Then load the "pieces" sprite, telling it to use the board.  When loading the
    'sprite piece, be sure to mention the number of the cell columns and rows in the
    'sprite.  If you don't, and then try to render the sprite using cells, nothing
    'will render at all.
    sprPiece.LoadImageSprite App.Path & "\Pieces.jpg", sprBoard.hdc, 6, 4
        
    'Now, just to show an animation example, we'll load an animatable sprite from
    'an old MegaMan game.  This one we'll put right on the buffer.
    sprSprite.LoadBMPSprite App.Path & "\Sprite.bmp", sprBuffer.hdc, 10, 1, RGB(0, 255, 0)
    
    'Our "runner" will be the same as the sprite.
    sprRunner.LoadBMPSprite App.Path & "\Sprite.bmp", sprBuffer.hdc, 10, 1, RGB(0, 255, 0)
    
    'Place the "runner" at the bottom of the screen.
    sprRunner.y = (Screen.Height / Screen.TwipsPerPixelY) - sprRunner.Cell_Height - 100
    
    'For the sake of example, draw a circle, line, and square to the buffer.
    'First set the "pen" (line).
    sprBuffer.SetPen RGB(128, 0, 0), 3, PS_SOLID
    'Then the "brush" (or fill)
    sprBuffer.BrushColor = RGB(0, 128, 0)
    'Then draw the circle.  Basically, a circle is treated as a rectangle.  Instead of
    'passing the center and radius, we pass the upper left, and lower right of the rectangle
    'that it would fit in.  This also allows us to make ellipses.
    sprBuffer.DrawCircle 100, 100, 200, 200
    'Change the pen color, and width.
    sprBuffer.PenColor = RGB(0, 0, 255)
    sprBuffer.PenWidth = 6
    'And draw a line.
    sprBuffer.DrawLine 210, 100, 310, 200
    'Once again, change the pen's color/width.
    sprBuffer.SetPen RGB(128, 128, 128), 2
    'And the brush.
    sprBuffer.BrushColor = RGB(128, 128, 0)
    'And draw the rectangle.
    sprBuffer.DrawRectangle 320, 100, 420, 200
    
    'Now, let's make our font a bit more interesting than the default.
    'We'll use Times New Roman, 24 pitch, at 310 degrees (counter-clockwise),
    'with 1000 as the "weight" (which is the heaviest "bold") and italics on.
    sprBuffer.ChangeFontStyle "Times New Roman", 24, 310, 1000, True
    'Make the color of the font red.
    sprBuffer.ChangeFontColor RGB(255, 0, 0)
    'And draw "Hello World!" on the screen.
    sprBuffer.DrawText "Hello World!", 100, 300
    
    'Since we're gonna stick in a tight "game" loop, we'll show the form now.
    Me.Show
           
    'Then, we'll use the ONE "piece" sprite to draw all 32 pieces on the board.  This should
    'give you a good idea of how you could use a single sprite to draw a tile-based world.
    '(I only used the loop and SetChessPiecePosition function to make things cleaner here.)
    'Note that here I turn off the "EnableErase" flag.  If this flag isn't set, then time
    'isn't wasted bitbltting the background to a "back storage buffer" before drawing
    'the sprite.  Here, not such a big deal, but in the game loop, it can make a huge
    'difference.  Keep in mind also that we're drawing this directly onto the board, and
    'not to the buffer.  This means we have to make sure to draw it onto the board
    'before drawing the board to the buffer.
    For lngCount = 0 To 31
    
        'Take a look at this function below:
        SetChessPiecePosition lngCount
        
        'Note that this will put a red dot in the upper left hand corner of EACH
        'cell since it's being drawn relative to the cells.  (Otherwise, it would
        'just put a dot in the upper left corner of the whole sprite.)  Note that
        'this is "destructive" in that it can't be undone unless you "get" the
        'pixel color first before changing it, and then change it back.
        sprPiece.SetPixelRGB 0, 0, 255, 0, 0, True
        sprPiece.DrawSprite , , , , , False
        
    Next lngCount
        
    'And finally we load the "transparent" square setting the sprite to transparent
    'by saying that the solid green colors on it are the transparency.  We'll also
    'draw this directly to the board.
    sprSquare.LoadBMPSprite App.Path & "\BlueSquare.bmp", sprBoard.hdc, , , RGB(0, 255, 0)
   
    'And last but not least, draw the "transparent" square.
    sprSquare.DrawSprite 95, 119, , , , False
    
    'Now let's draw our board...
    'Note that since we won't be animating any of these, I'm turning off the
    '"EnableErase" flag.  This cuts your draw-time in half.  This board will now
    'draw directly to the buffer.  (The math is to put it right in the middle of the screen.)
    sprBoard.DrawSprite (sprBuffer.Width / 2) - (sprBoard.Width / 2), (sprBuffer.Height / 2) - (sprBoard.Height / 2), , , , False

    'Now we'll draw our animated sprite, but since we'll be animating it, we'll leave
    'the "EnableErase" flag set to "True" (which is the default).
    sprSprite.SetCell 1, 1
    sprSprite.DrawSprite
    
    'Snag the current tick counts before the run loop starts (so we can compare it
    'later to get the Frames Per Second (FPS)  We add 1000 milliseconds (or 1 second) to
    'the current "TickCount" (which is the number of milliseconds since the computer was
    'booted) so we can get the number of frames rendered every second.
    lngTick = GetBetterTick + 1000
    
    While booIsRunning
        RenderScene 'See the following sub.
        
        'To calculate the total framerate.  (Will be printed in the caption.)
        lngFPS = lngFPS + 1
        If lngTick < GetBetterTick Then
            Me.Caption = "FPS: " & CStr(lngFPS)
            lngTick = GetBetterTick + 1000
            lngFPS = 0
        End If
        
        
        
        'Always use DoEvents when running a tight loop like this.  Otherwise, your
        'keypresses and such will never register.
        DoEvents
        
        'I like to throw this in there too.  It will seem to "destroy" your framerates
        'because it will ALWAYS drop your framerates to a certain extent.
        'However, I find that it allows windows to "rest" the correct amount of time
        'to keep it from "sticking."  To see what I mean, run your task manager, and
        'keep it in the background while you run this.  Note that your CPU usage stays at
        '100%.  This is because you NEVER let it rest.  Then press "s" to turn on the
        'following sleep function.  Even though you're sleeping for only 1 millisecond
        '(possibly a few more depending on the OS) the amount of break it gives the
        'CPU is amazing.  Without Sleep on, on my PC (ATI RADEON 9600 Pro @ 1024x768), it's around
        'running at roughly 530 FPS.  Keep in mind that since it's "seems" to slow down to
        'wait for the Sleep, you're really not losing a whole lot of time by using it.
        'The monitor will only REFRESH at the VSync speed ANYWAY.  If your framerates in
        'your final game are above 60 FPS, it doesn't hurt to use it, and give the CPU a
        'break.  With this turned on, I get around 400 FPS.
        If booUseSleep Then Sleep 1
        
    Wend
    
    'Once we exit this loop, we know the "game" has ended.
    
    'Show the mouse again.
    While ShowCursor(1) < 0: Wend
    
    'And close the form.
    Unload Me
End Sub

Private Sub RenderScene()
    'Keep in mind that you're going to have a LOT more going on with a render than
    'what I'm doing here, so your framerates will show that.  For now, I'm simply
    'going to "erase" the sprite, then draw it again.  When using more than one
    'sprite you need to remember to erase them in the REVERSE order that you draw
    'them.  Otherwise, you'll run into trouble.
    
    'Just so you can actually see the animation, we'll slow down how often we change
    'the column of the sprite we're using.  The .Tag_Long is a handy place to store the
    'amount of time we're going to wait.  Note that this also sets the speed that the
    'sprite cycles to a constant, regardless of framerate.  (Unless the framerate
    'drops enough that every frame the animation is rendered.)
    With sprSprite
        Debug.Print GetBetterTick
        If .Tag_Long < GetBetterTick Then
            .Tag_Long = GetBetterTick + 75
            .Column = .Column + 1
            If .Column > 10 Then .Column = 1
        End If
    End With
    
    'Keep in mind that this draws to the sprBuffer.  First we call the EraseSprite
    'procedure.  This will clear the LAST PLACE the sprite was drawn.  Keep that in
    'mind. If you call the DrawSprite twice before an EraseSprite, it only will
    'erase the LAST one.  And also know that you ALWAYS ALWAYS ALWAYS erase in the
    'OPPOSITE order you draw.
    sprRunner.EraseSprite
    sprSprite.EraseSprite
    
    'Then draw the sprite again.  Since the mouse-move event updates the position of this
    'sprite, all we have to do here is just draw it.
    sprSprite.DrawSprite
    
    'As far as animation goes, do the same for the runner.
    If sprRunner.Tag_Long < GetBetterTick Then
        With sprRunner
            .Tag_Long = GetBetterTick + 75
            .Column = .Column + 1
            If .Column > 10 Then .Column = 1
            'But add the movement code in here.
            .x = .x - 7
            'If he's passed the left side of the screen, send him back to the right.
            If .x < -.Cell_Width Then
                .x = (Screen.Width / Screen.TwipsPerPixelX)
            End If
        End With
    End If
    
    'Since the runner is drawn AFTER the mouse-sprite, he will appear "on top" of it.
    sprRunner.DrawSprite
    
    'So you have to call the sprBuffer.DrawSprite to get it to the main form.  The reason
    'for this is to keep "flicker" down.  You should only call an update to the main form
    'once EVERYTHING has been updated.  (Which, right now, means just the two sprites.)
    'Obviously, you have no reason to save the background of the sprBuffer, so don't
    'EnableErase.
    sprBuffer.DrawSprite , , , , , False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    booIsRunning = False
    
    'Always destroy what you create!
    Set sprRunner = Nothing
    Set sprSleep = Nothing
    Set sprSprite = Nothing
    Set sprPiece = Nothing
    Set sprBoard = Nothing
    Set sprSquare = Nothing
    Set sprBuffer = Nothing
End Sub

Private Sub SetChessPiecePosition(ByVal PieceNumber As Long)
    'This is not how you would actually pull something like this off in a real game, but
    'it will work for this demo.  Here, we actually move the ONE sprite to each of the
    '32 positions, then tell it which cell/column the image is in that we want to draw.
    'By doing this we use just one sprite object to draw all 32 chess pieces on the board.
    
    'There's a number of different ways to change this data.  For instance:
    Select Case PieceNumber
        'You can change the values individually.
        Case 0
            sprPiece.x = 20
            sprPiece.y = 6
            sprPiece.Column = 1
            sprPiece.Row = 2
            
        'Or change the position/cell as a group.
        Case 1
            sprPiece.SetPosition 39, 6
            sprPiece.SetCell 5, 2
            
        'Or change it all with one call.
        Case 2
            sprPiece.SetPosAndCell 58, 6, 1, 1
        Case 3
            sprPiece.SetPosAndCell 77, 6, 5, 1
        Case 4
            sprPiece.SetPosAndCell 96, 6, 3, 1
        Case 5
            sprPiece.SetPosAndCell 115, 6, 4, 1
        Case 6
            sprPiece.SetPosAndCell 134, 6, 2, 2
        Case 7
            sprPiece.SetPosAndCell 153, 6, 4, 2
        Case 8
            sprPiece.SetPosAndCell 20, 25, 6, 2
        Case 9
            sprPiece.SetPosAndCell 39, 25, 3, 2
        Case 10
            sprPiece.SetPosAndCell 58, 25, 6, 2
        Case 11
            sprPiece.SetPosAndCell 77, 25, 3, 2
        Case 12
            sprPiece.SetPosAndCell 96, 25, 6, 2
        Case 13
            sprPiece.SetPosAndCell 115, 25, 3, 2
        Case 14
            sprPiece.SetPosAndCell 134, 25, 6, 2
        Case 15
            sprPiece.SetPosAndCell 153, 25, 3, 2
        Case 16
            sprPiece.SetPosAndCell 20, 120, 3, 4
        Case 17
            sprPiece.SetPosAndCell 39, 120, 6, 4
        Case 18
            sprPiece.SetPosAndCell 58, 120, 3, 4
        Case 19
            sprPiece.SetPosAndCell 77, 120, 6, 4
        Case 20
            sprPiece.SetPosAndCell 96, 120, 3, 4
        Case 21
            sprPiece.SetPosAndCell 115, 120, 6, 4
        Case 22
            sprPiece.SetPosAndCell 134, 120, 3, 4
        Case 23
            sprPiece.SetPosAndCell 153, 120, 6, 4
        Case 24
            sprPiece.SetPosAndCell 20, 139, 4, 4
        Case 25
            sprPiece.SetPosAndCell 39, 139, 2, 4
        Case 26
            sprPiece.SetPosAndCell 58, 139, 4, 3
        Case 27
            sprPiece.SetPosAndCell 77, 139, 2, 3
        Case 28
            sprPiece.SetPosAndCell 96, 139, 6, 3
        Case 29
            sprPiece.SetPosAndCell 115, 139, 1, 3
        Case 30
            sprPiece.SetPosAndCell 134, 139, 5, 4
        Case 31
            sprPiece.SetPosAndCell 153, 139, 1, 4
    End Select
End Sub

'Since GetTickCount returns an "unsigned" long, if your PC has been running long
'enough (24.85 days), it'll actually return a NEGATIVE value.  This keeps that from
'happening by keeping track of when your PC was started, and adjusting for it.  After
'49.7 days, the tick count is back to 0 and all is kosher.  This will return the time
'elapsed since the last time the function was "reset."
Private Function GetBetterTick() As Long
    Static LastTime As Long
    If LastTime >= 0 And GetTickCount < 0 Then LastTime = GetTickCount
    If LastTime <= 0 And GetTickCount > 0 Then LastTime = GetTickCount
    GetBetterTick = GetTickCount - LastTime
End Function
