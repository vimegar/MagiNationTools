Attribute VB_Name = "modViewport"
Option Explicit

Public Const VIEWPORT_TILE_WIDTH = 10
Public Const VIEWPORT_TILE_HEIGHT = 9

Public Const VIEWPORT_ZOOM = 2

Public Type tViewport
    xTile As Integer
    yTile As Integer
End Type

Public gViewport As tViewport
Public gBackBuffer As Object

Public Sub DrawViewportFromDC(SrcDC As Long)

    BitBlt gBackBuffer.hDC, 0, 0, (VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH), (VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT), SrcDC, Int(gViewport.xPixel), Int(gViewport.yPixel), SRCCOPY
    
End Sub

Public Sub DrawViewportToDC(DestDC As Long)

    StretchBlt DestDC, 0, 0, (VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH) * VIEWPORT_ZOOM, (VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT) * VIEWPORT_ZOOM, gBackBuffer.hDC, 0, 0, (VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH), (VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT), SRCCOPY
    
End Sub


Public Sub InitViewport()

    Set gBackBuffer = frmMain.picBackBuffer

    gBackBuffer.Width = (VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH)
    gBackBuffer.Height = (VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT)

End Sub


