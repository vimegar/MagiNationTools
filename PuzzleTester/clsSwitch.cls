VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnXInTiles As Integer
Private mnYInTiles As Integer

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

