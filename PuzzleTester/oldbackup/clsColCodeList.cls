VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnTileID() As Integer
Private mnTileCount As Integer
Private mbWalkable As Boolean

Public Property Get TileCount() As Integer

    TileCount = mnTileCount

End Property


Public Property Get TileID(Index As Integer) As Integer

    TileID = mnTileID(Index)

End Property

Public Property Let TileID(Index As Integer, nNewValue As Integer)

    mnTileID(Index) = nNewValue

End Property

Public Sub AddTileToList(TileID As Integer)

    mnTileCount = mnTileCount + 1
    ReDim Preserve mnTileID(mnTileCount - 1)
    
    mnTileID(mnTileCount - 1) = TileID

End Sub
Public Property Let Walkable(bNewValue As Boolean)

    mbWalkable = bNewValue

End Property

Public Property Get Walkable() As Boolean

    Walkable = mbWalkable

End Property

