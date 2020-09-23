Attribute VB_Name = "modDD"
Option Explicit

'The Main Direct Draw Object
Public ddMain As DirectDraw7

'This is the primary buffer and its description
Public PrimBuf As DirectDrawSurface7
Public PrimBufDesc As DDSURFACEDESC2

'This is the back buffer and its description
Public BackBuf As DirectDrawSurface7
Public BackBufDesc As DDSURFACEDESC2

'This is used in the creation of the back buffer
Public Caps As DDSCAPS2

'The is the Surface we put the background picture in
Public BGSurf As DirectDrawSurface7
'This is used for the blitting process
Public rBGSurf As RECT

'The is the Surface we put the sprites picture in
Public SpriteSurf As DirectDrawSurface7
'This is used for the blitting process
Public rSpriteSurf As RECT

'This creates the Primary Surface and The Backbuffer
Sub DD_CreatePrimBackBuf()

    'These are used to make the primary surface
    PrimBufDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    PrimBufDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    'This backbuffer count can be changed to 2 for triple buffering
    PrimBufDesc.lBackBufferCount = 1
    Set PrimBuf = ddMain.CreateSurface(PrimBufDesc)

    'These are used to make the backbuffer
    Caps.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuf = PrimBuf.GetAttachedSurface(Caps)
    Call BackBuf.GetSurfaceDesc(PrimBufDesc)

    'This sets the Fore color of the text at the top of the form
    Call BackBuf.SetForeColor(RGB(150, 0, 0))
    'This sets the back color of the text. This isn't really needed but, what the hell
    Call BackBuf.SetFontBackColor(vbWhite)
    'This sets the font to be transparent. If it was set to false a white block would
    'show up around the text if you moved the sprite under it
    Call BackBuf.SetFontTransparency(True)

End Sub

'This is the sub that blits the surfaces to the backbuffer
Sub DD_BltFast(rTop As Integer, rLeft As Integer, Width As Integer, Height As Integer, Surface As DirectDrawSurface7, srcRect As RECT, X As Integer, Y As Integer, Transparency As Boolean)

    'this is used for the image sizing process
    'please note that my sprite is actually 200 pixels wide but it is set to 50 pixels
    'in the main rendering loop
    srcRect.Top = rTop
    srcRect.Left = rLeft
    srcRect.Right = Width
    srcRect.Bottom = Height

    'If there is no transparency then
    If Transparency = False Then
        'blit it with no transparency
        Call BackBuf.BltFast(X, Y, Surface, srcRect, DDBLTFAST_WAIT)
    'If there is transparency then
    Else
        'blit it with transparency
        Call BackBuf.BltFast(X, Y, Surface, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

End Sub

'This creates the surfaces from their files
Sub DD_CreateSurfFromFile(FileName As String, Surface As DirectDrawSurface7, SurfDesc As DDSURFACEDESC2, Width As Long, Height As Long)

    'This is part of the creation process
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'This sets the surfaces description to an off screen plain
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'This sets the height of the surfaces description
    SurfDesc.lHeight = Height
    'This sets the width of the surfaces description
    SurfDesc.lWidth = Width

    'This sets the surface to the file with the above settings
    Set Surface = ddMain.CreateSurfaceFromFile(FileName, SurfDesc)

    'These are used for transparency during the blitting process
    Dim ColorKey As DDCOLORKEY
    'the transparent color will be black
    ColorKey.high = 0
    ColorKey.low = 0
    'This sets the surfaces color key to the color keys above
    Call Surface.SetColorKey(DDCKEY_SRCBLT, ColorKey)

End Sub

'This sub basically calls the above sub for easier transportation to the main initialization
'This makes it easier because I have multiple files I want in the application
Sub DD_CreateGraphicsFromFile()
    Call DD_CreateSurfFromFile(App.Path & "\bg.bmp", BGSurf, BackBufDesc, 640, 480)
    Call DD_CreateSurfFromFile(App.Path & "\sprite.bmp", SpriteSurf, BackBufDesc, 200, 41)
End Sub
