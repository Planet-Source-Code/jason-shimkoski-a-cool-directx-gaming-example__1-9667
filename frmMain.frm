VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "First Game Thing"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    'this ends the app if the user clicks the form
    DX_EndIt
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'these are the keys used for moving the sprite.
    Select Case KeyCode
    Case vbKeyUp
        'This moves the sprites Y-Coordinate up
        sY = sY - 10
    Case vbKeyDown
        'This moves the sprites Y-Coordinate down
        sY = sY + 10
    Case vbKeyLeft
        'This moves the sprites X-Coordinate left
        sX = sX - 10
    Case vbKeyRight
        'This moves the sprites X-Coordinate right
        sX = sX + 10
    'if the user presses Escape it will end the program
    Case vbKeyEscape
        DX_EndIt
    End Select

End Sub

Private Sub Form_Load()
    'this goes to the main initialization sub
    DX_Init
End Sub
