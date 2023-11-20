Attribute VB_Name = "modGeneral"
Option Explicit

Public glMainHDC As Long

Public Sub CreateMask(SrcDC As Long, SrcPixelWidth As Integer, SrcPixelHeight As Integer, MaskDC As Long)

    Dim x As Integer
    Dim y As Integer

    For y = 0 To (SrcPixelHeight - 1)
        For x = 0 To (SrcPixelWidth - 1)
            If GetPixel(SrcDC, x, y) = vbBlack Then
                SetPixel MaskDC, x, y, vbWhite
            Else
                SetPixel MaskDC, x, y, vbBlack
            End If
        Next x
    Next y

End Sub

Public Sub GetDetectTile(Player As tPlayer, bLeft As Boolean, bUp As Boolean, bRight As Boolean, bDown As Boolean, ScrollRate As Single, xDetect As Integer, yDetect As Integer, xPixelNew As Single, yPixelNew As Single)

    xPixelNew = Player.xTile - (ScrollRate * bRight) + (ScrollRate * bLeft)
    yPixelNew = Player.yTile - (ScrollRate * bDown) + (ScrollRate * bUp)
    
    xDetect = (((xPixelNew + LEFT_PIXEL_DETECT) * -bLeft) + ((xPixelNew + RIGHT_PIXEL_DETECT) * -(Not bLeft))) \ TILE_PIXEL_WIDTH
    yDetect = (((yPixelNew + UP_PIXEL_DETECT) * -bUp) + ((yPixelNew + DOWN_PIXEL_DETECT) * -(Not bUp))) \ TILE_PIXEL_HEIGHT
    
End Sub

Public Sub Main()

    frmMain.Show

    InitTiles
    InitPlayer
    InitColCodes
    InitMaps
    
    SelectMap UNDGEYSER11
    
End Sub

Public Function KeyDown(KeyCode As Integer) As Boolean

    KeyDown = GetAsyncKeyState(CLng(KeyCode))

End Function

Public Sub UpdateScreen(DestDC As Long)

    gMapBuffer.Cls

    DrawMap
    DrawPlayer gMapBuffer.hDC

    BitBlt frmMain.picDisplay.hDC, 0, 0, 320, 288, gMapBuffer.hDC, 0, 0, SRCCOPY

End Sub


