VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCTArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCompareTiles() As clsCT
Private mnTileCount As Integer

Public Property Get TileCount() As Integer

    TileCount = mnTileCount

End Property

Public Property Let TileCount(nNewValue As Integer)

    mnTileCount = nNewValue
    ReDim Preserve moCompareTiles(mnTileCount)

End Property

Public Property Set CompareTiles(Index As Integer, oNewValue As clsCT)

    Set moCompareTiles(Index) = oNewValue

End Property

Public Property Get CompareTiles(Index As Integer) As clsCT

    Set CompareTiles = moCompareTiles(Index)

End Property
