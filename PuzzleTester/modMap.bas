Attribute VB_Name = "modMap"
Option Explicit

Public Const NUM_MAPS = 11

Public Const TILE_WIDTH_IN_PIXELS = 16
Public Const TILE_HEIGHT_IN_PIXELS = 16

Public Const DISPLAY_WIDTH_IN_TILES = 10
Public Const DISPLAY_HEIGHT_IN_TILES = 9

Public Enum MAP_NAME_CONSTANTS
    UNDGEYSER01 = 0
    UNDGEYSER02 = 1
    UNDGEYSER03 = 2
    UNDGEYSER04 = 3
    UNDGEYSER05 = 4
    UNDGEYSER06 = 5
    UNDGEYSER07 = 6
    UNDGEYSER08 = 7
    UNDGEYSER09 = 8
    UNDGEYSER10 = 9
    UNDGEYSER11 = 10
End Enum

Public gMaps(NUM_MAPS - 1) As clsMap

Private mtCurrentMap As MAP_NAME_CONSTANTS

Public Property Set CurrentMap(oNewValue As clsMap)

    Set gMaps(mtCurrentMap) = oNewValue

End Property

Public Property Get CurrentMap() As clsMap

    Set CurrentMap = gMaps(mtCurrentMap)

End Property

Public Sub DrawMap(BufferDC As Long)

    On Error Resume Next

    Dim x As Integer
    Dim y As Integer
    
    For y = -4 To 4
        For x = -4 To 5
            BitBlt BufferDC, (x + 4) * TILE_WIDTH_IN_PIXELS, (y + 4) * TILE_HEIGHT_IN_PIXELS, TILE_WIDTH_IN_PIXELS, TILE_HEIGHT_IN_PIXELS, gTileBitmaps(CurrentMap.MapData(gPlayer.xInTiles + x, gPlayer.yInTiles + y)).hdc, 0, 0, vbSrcCopy
        Next x
    Next y

End Sub

Public Sub InitMap()

'UNDGEYSER10
    Set gMaps(UNDGEYSER10) = New clsMap
    gMaps(UNDGEYSER10).MapName = UNDGEYSER10
    gMaps(UNDGEYSER10).LoadMap App.Path & "\Maps\SCR_UNDGEYSER10.map"
    InitHotspot UNDGEYSER10
    InitSwitch UNDGEYSER10

'UNDGEYSER11
    Set gMaps(UNDGEYSER11) = New clsMap
    gMaps(UNDGEYSER11).MapName = UNDGEYSER11
    gMaps(UNDGEYSER11).LoadMap App.Path & "\Maps\SCR_UNDGEYSER11.map"
    InitHotspot UNDGEYSER11
    InitSwitch UNDGEYSER11

End Sub

Public Sub SelectMap(MapName As MAP_NAME_CONSTANTS)

    mtCurrentMap = MapName

End Sub


