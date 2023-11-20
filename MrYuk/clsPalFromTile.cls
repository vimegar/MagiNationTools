VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPalFromTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moFeet(4) As clsPool
Private mnFeetIndex(4) As Integer
Private mnMatchCount As Integer
Private mnColorCount As Integer
Private mnSourceX As Integer
Private mnSourceY As Integer

Public Property Let ColorCount(nNewValue As Integer)

    mnColorCount = nNewValue
    
End Property

Public Property Let FeetIndex(Index As Integer, nNewValue As Integer)

    mnFeetIndex(Index) = nNewValue

End Property

Public Property Get FeetIndex(Index As Integer) As Integer

    FeetIndex = mnFeetIndex(Index)

End Property

Public Property Let SourceY(nNewValue As Integer)

    mnSourceY = nNewValue

End Property

Public Property Get SourceY() As Integer

    SourceY = mnSourceY

End Property

Public Property Let SourceX(nNewValue As Integer)

    mnSourceX = nNewValue

End Property

Public Property Get SourceX() As Integer

    SourceX = mnSourceX

End Property

Public Property Get Feet(Index As Integer) As clsPool

    Set Feet = moFeet(Index)

End Property

Public Property Set Feet(Index As Integer, oNewValue As clsPool)

    Set moFeet(Index) = oNewValue

End Property

Public Property Get MatchCount() As Integer

    MatchCount = mnMatchCount

End Property

Public Property Let MatchCount(nNewValue As Integer)

    mnMatchCount = nNewValue

End Property

Public Property Get ColorCount() As Integer

    ColorCount = mnColorCount

End Property


