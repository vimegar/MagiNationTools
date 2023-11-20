VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTileGetterError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnX As Integer
Private mnY As Integer
Private mnColorCount As Integer
Private mnType As TILE_GETTER_ERROR_TYPES

Public Property Let nType(nNewValue As TILE_GETTER_ERROR_TYPES)

    mnType = nNewValue

End Property

Public Property Get nType() As TILE_GETTER_ERROR_TYPES

    nType = mnType

End Property

Public Property Get ColorCount() As Integer

    ColorCount = mnColorCount

End Property

Public Property Let ColorCount(nNewValue As Integer)

    mnColorCount = nNewValue

End Property

Public Property Get X() As Integer

    X = mnX

End Property

Public Property Let X(nNewValue As Integer)

    mnX = nNewValue

End Property

Public Property Get Y() As Integer

    Y = mnY

End Property

Public Property Let Y(nNewValue As Integer)

    mnY = nNewValue

End Property

