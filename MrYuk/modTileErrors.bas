Attribute VB_Name = "modTileErrors"
Option Explicit

Public Enum TILE_GETTER_ERROR_TYPES
    TooManyColors = 0
    NoPixelMatch = 1
    NoPaletteMatch = 2
End Enum

Private moTileGetterErrors() As clsTileGetterError
Private mnErrorCount As Integer

Public Sub AddTileGetterError()

    If mnErrorCount = 0 Then
        ReDim moTileGetterErrors(0)
    End If
    
    mnErrorCount = mnErrorCount + 1
    ReDim Preserve moTileGetterErrors(mnErrorCount)
    
    Set moTileGetterErrors(mnErrorCount) = New clsTileGetterError

End Sub


Public Property Let TileGetterErrorCount(nNewValue As Integer)

    mnErrorCount = nNewValue

End Property

Public Property Get TileGetterErrorCount() As Integer

    TileGetterErrorCount = mnErrorCount

End Property


Public Property Set TileGetterErrors(Index As Integer, oNewValue As clsTileGetterError)

    Set moTileGetterErrors(Index) = oNewValue

End Property

Public Property Get TileGetterErrors(Index As Integer) As clsTileGetterError

    Set TileGetterErrors = moTileGetterErrors(Index)

End Property


