VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpriteTileGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnTotalTiles As Integer
Private msName As String
Private mtTiles() As New clsSpriteTile

Public Sub AddTile()

'***************************************************************************
'   Add an entry into the sprite group
'***************************************************************************

    ReDim Preserve mtTiles(UBound(mtTiles) + 1)
    
End Sub


Public Sub DeleteTile(Index As Integer)

'***************************************************************************
'   Delete a sprite group entry based on its index
'***************************************************************************

    Dim i As Integer
    
    For i = Index To UBound(mtTiles) - 1
        Set mtTiles(i) = mtTiles(i + 1)
    Next i
    
    ReDim Preserve mtTiles(UBound(mtTiles) - 1)

End Sub

Public Property Let Name(sNewValue As String)

    msName = sNewValue

End Property

Public Property Get Name() As String

    Name = msName

End Property

Public Property Get TileCount() As Integer

    TileCount = UBound(mtTiles)

End Property

Public Property Set Tiles(Index As Integer, oNewValue As clsSpriteTile)

    Set mtTiles(Index) = oNewValue

End Property

Public Property Get Tiles(Index As Integer)

    Set Tiles = mtTiles(Index)

End Property

Private Sub Class_Initialize()

    ReDim mtTiles(0)

End Sub


