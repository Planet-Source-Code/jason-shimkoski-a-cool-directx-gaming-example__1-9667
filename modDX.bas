Attribute VB_Name = "modDX"
Option Explicit

'This is a very basic program to help get people started in creating games.
'This has no sort of collision detection what-so-ever and I know about the end of
'the screen bug where the sprite just disappears. This also doesn't use Direct
'Input for the keystrokes. However, this is a good way of learning the basics
'of Direct Draw and Direct Sound despite the ugly sprite.
'
'Thanks,
'Jason Shimkoski (basspler@aol.com)


'The Main DirectX Object
Public dxMain As New DirectX7

'The Sprites Current X and Y Values
Public sX As Integer
Public sY As Integer

'Checks to See if The Main Loop should stop or not
Public running As Boolean

'This sets the screens display and the cooperative levels
Sub DX_SetDisplayCoopLevel(Hdl As Long, sWidth As Long, sHeight As Long, sBPP As Long)
    'This one is for the Direct Sound
    Call dsMain.SetCooperativeLevel(Hdl, DSSCL_PRIORITY)
    'This one is for Direct Draw
    Call ddMain.SetCooperativeLevel(Hdl, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX)
    'This sets the display mode
    Call ddMain.SetDisplayMode(sWidth, sHeight, sBPP, 0, DDSDM_DEFAULT)
End Sub

'This restores everything to normal when the Program is exited
Sub DX_RestoreDisplayCoopLevel(Hdl As Long)
    'This sets the Direct Sound Cooperative Level to normal
    Call dsMain.SetCooperativeLevel(Hdl, DSSCL_NORMAL)
    'This sets the Direct Draw Cooperative Level to normal
    Call ddMain.SetCooperativeLevel(Hdl, DDSCL_NORMAL)
    'This restores the users default Display Mode
    Call ddMain.RestoreDisplayMode
End Sub

'This is where everything comes together
Sub DX_Init()

    'This tells the loop below that it is running
    running = True

    On Error Resume Next
    'This creates the Direct Draw object
    Set ddMain = dxMain.DirectDrawCreate("")
    'This creates the Direct Sound object
    Set dsMain = dxMain.DirectSoundCreate("")
    
    'This calls the sub from above that sets the Display and the Cooperative Levels
    Call DX_SetDisplayCoopLevel(frmMain.hWnd, 640, 480, 16)
    'This calls a sub from modDD that creates the Primary Surface and the Backbuffer
    DD_CreatePrimBackBuf

    'This calls a sub from modDD that Creates Graphics from their Files
    DD_CreateGraphicsFromFile
    'This calls a sub from modDS that Creates Sounds from their Files
    DS_CreateSoundsFromFile

    'This calls a sub from modDS that Starts the Sound file and tells whether
    'it should be looped or not
    Call DS_PlaySound(False)

    'This is the main render loop
    Do
        'This blits a black color fill to the back buffer
        Call BackBuf.BltColorFill(rBGSurf, RGB(0, 0, 0))
        'This calls a sub from modDD that blits the Background to the back buffer
        Call DD_BltFast(0, 0, 640, 480, BGSurf, rBGSurf, 0, 0, False)
        'This calls a sub from modDD that blits the Sprite to the back buffer
        Call DD_BltFast(0, 0, 50, 41, SpriteSurf, rSpriteSurf, sX, sY, True)
        'This draws the text at the top of the form
        'please note that if this was created before the sprite, the sprite wouldn't
        'be able to cross behind it
        Call BackBuf.DrawText(200, 25, "This is an Awesome Thing, Dude!", False)
        'This flips everything to the Primary Surface
        Call PrimBuf.Flip(Nothing, DDFLIP_WAIT)
        'This is so windows can process other events
        DoEvents
    'This tells the program to loop until the user quits the program
    Loop Until running = False

End Sub

'This ends the app
Sub DX_EndIt()
    'This tells the render loop to stop working
    running = False
    'This calls a sub from above that sets everything to normal
    Call DX_RestoreDisplayCoopLevel(frmMain.hWnd)
    'This ends the application
    End
End Sub
