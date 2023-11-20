Attribute VB_Name = "modColCodes"
Option Explicit

Public Const NUM_COL_CODES = 5

Public Enum COL_CODES
    COL_GROUND = 0
    COL_WALL = 1
    COL_SWITCH = 2
    COL_HOTSPOT = 3
    COL_MAPLINK = 4
End Enum

Public gColCodeLists(NUM_COL_CODES - 1) As clsColCodeList

Public Sub HitDetect(Player As tPlayer, bLeft As Boolean, bUp As Boolean, bRight As Boolean, bDown As Boolean, ScrollRate As Single)
    
    
End Sub


Public Sub InitColCodes()

    'GROUND
    Set gColCodeLists(0) = New clsColCodeList
    gColCodeLists(0).Walkable = True
    gColCodeLists(0).AddTileToList 1
    gColCodeLists(0).AddTileToList 11
    gColCodeLists(0).AddTileToList 12
    gColCodeLists(0).AddTileToList 13
    gColCodeLists(0).AddTileToList 14
    
    'WALL
    Set gColCodeLists(1) = New clsColCodeList
    gColCodeLists(1).Walkable = False
    gColCodeLists(1).AddTileToList 0
    gColCodeLists(1).AddTileToList 2
    gColCodeLists(1).AddTileToList 3
    gColCodeLists(1).AddTileToList 4
    gColCodeLists(1).AddTileToList 5
    
    'SWITCH
    Set gColCodeLists(2) = New clsColCodeList
    gColCodeLists(2).Walkable = False
    gColCodeLists(2).AddTileToList 0
    gColCodeLists(2).AddTileToList 2
    gColCodeLists(2).AddTileToList 3
    gColCodeLists(2).AddTileToList 4
    gColCodeLists(2).AddTileToList 5
    
    'HOTSPOT
    Set gColCodeLists(3) = New clsColCodeList
    gColCodeLists(3).Walkable = True
    gColCodeLists(3).AddTileToList 10
    
    'MAPLINK
    Set gColCodeLists(4) = New clsColCodeList
    gColCodeLists(4).Walkable = False

End Sub


