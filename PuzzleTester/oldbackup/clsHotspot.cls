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

Private mnXTile As Integer
Private mnYTile As Integer
Private mnCode As Integer

Public Property Let nCode(nNewValue As Integer)

    mnCode = nNewValue

End Property

Public Property Get nCode() As Integer

    nCode = mnCode

End Property

Public Property Let xTile(nNewValue As Integer)

    mnXTile = nNewValue

End Property

Public Property Get xTile() As Integer
    
    xTile = mnXTile

End Property


Public Property Let yTile(nNewValue As Integer)

    mnYTile = nNewValue

End Property


Public Property Get yTile() As Integer

    yTile = mnYTile

End Property

