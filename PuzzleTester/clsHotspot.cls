VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHotspot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnXInTiles As Integer
Private mnYInTiles As Integer
Private mtDirection As DIRECTION_CONSTANTS

Public Property Let Direction(tNewValue As DIRECTION_CONSTANTS)

    mtDirection = tNewValue

End Property

Public Property Get Direction() As DIRECTION_CONSTANTS

    Direction = mtDirection

End Property

Public Property Get xInTiles() As Integer

    xInTiles = mnXInTiles

End Property

Public Property Let xInTiles(nNewValue As Integer)

    mnXInTiles = nNewValue

End Property

Public Property Get yInTiles() As Integer

    yInTiles = mnYInTiles

End Property

Public Property Let yInTiles(nNewValue As Integer)

    mnYInTiles = nNewValue

End Property
