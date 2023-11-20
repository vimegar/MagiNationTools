Attribute VB_Name = "modMain"
Option Explicit

Public gBuffer As Object
Public gDisplay As Object
Public Sub Main()

    Set gBuffer = frmMain.picBuffer
    Set gDisplay = frmMain.picDisplay

    InitColCodes
    InitPlayer
    InitTiles
    InitMap
    
    InitGeyser2
    
    SelectMap UNDGEYSER11
    
    frmMain.Show
    UpdateScreen

End Sub

Public Sub UpdateScreen()

    gBuffer.Cls
    
    DrawMap gBuffer.hdc
    DrawPlayer gBuffer.hdc
    
    StretchBlt gDisplay.hdc, 0, 0, gDisplay.Width, gDisplay.Height, gBuffer.hdc, 0, 0, gBuffer.Width, gBuffer.Height, vbSrcCopy
    gDisplay.Refresh

End Sub

Public Function KeyDown(KeyCode As Integer) As Boolean

    KeyDown = GetAsyncKeyState(CLng(KeyCode))

End Function

