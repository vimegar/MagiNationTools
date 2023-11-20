Attribute VB_Name = "modMaps"
Option Explicit

Public Const TILE_PIXEL_WIDTH = 16
Public Const TILE_PIXEL_HEIGHT = 16

Public Const NUM_MAPS = 11

Public Enum MAP_NAME_CONSTANTS
    UNDGEYSER10 = 9
    UNDGEYSER11 = 10
End Enum

Public gMaps(NUM_MAPS - 1) As clsMap
Public gnCurrentMap As MAP_NAME_CONSTANTS

Public gMapBuffer As Object



Public Sub InitMaps()

    Set gMapBuffer = frmMain.picMapBuffer

    Dim i As Integer
    For i = 0 To (NUM_MAPS - 1)
        Set gMaps(i) = New clsMap
    Next i

    Dim sPath As String
    sPath = App.Path & "\Maps\"

    gMaps(9!).LoadMap sPath & "SCR_UNDGEYSER10.map"
    gMaps(10).LoadMap sPath & "SCR_UNDGEYSER11.map"

    For i = 0 To (NUM_MAPS - 1)
        gMaps(i).Index = i
        gMaps(i).InitHotspots
    Next i

    gnCurrentMap = UNDGEYSER11

End Sub

Public Sub SelectMap(Index As MAP_NAME_CONSTANTS)

    gnCurrentMap = Index

    gMapBuffer.Width = (gMaps(gnCurrentMap).TileWidth * TILE_PIXEL_WIDTH)
    gMapBuffer.Height = (gMaps(gnCurrentMap).TileHeight * TILE_PIXEL_HEIGHT)
    
    DrawMap

End Sub


Public Sub DrawMap()

    Dim xTile As Integer
    Dim yTile As Integer
    Dim dx As Integer
    Dim dy As Integer
    
    'On Error GoTo HandleErrors
    '
    'For yTile = 0 To (gMaps(gnCurrentMap).TileHeight - 1)
    '    For xTile = 0 To (gMaps(gnCurrentMap).TileWidth - 1)
    '        BitBlt gMapBuffer.hDC, (xTile * TILE_PIXEL_WIDTH), (yTile * TILE_PIXEL_HEIGHT), TILE_PIXEL_WIDTH, TILE_PIXEL_HEIGHT, gTileBitmaps(gMaps(gnCurrentMap).MapData(xTile, yTile)).hDC, 0, 0, SRCCOPY
    '    Next xTile
    'Next yTile
    
    On Error Resume Next
    
    For yTile = -5 To 5
        For xTile = -5 To 5
            dx = (xTile * TILE_PIXEL_WIDTH) + (gViewport.xPixel + ((VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH) \ 2))
            dy = (yTile * TILE_PIXEL_HEIGHT) + (gViewport.yPixel + ((VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT) \ 2))
            BitBlt gMapBuffer.hDC, dx, dy, TILE_PIXEL_WIDTH, TILE_PIXEL_HEIGHT, gTileBitmaps(gMaps(gnCurrentMap).MapData(dx \ TILE_PIXEL_WIDTH, dy \ TILE_PIXEL_HEIGHT)).hDC, 0, 0, SRCCOPY
        Next xTile
    Next yTile
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modMaps:DrawMap Error"
End Sub

