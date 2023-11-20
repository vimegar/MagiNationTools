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
Private moDependency As clsPalFromPal
Private mnMatchCount As Integer

Public Property Get Feet(Index As Integer) As clsPool

    Set Feet = moFeet(Index)

End Property

Public Property Set Feet(Index As Integer, oNewValue As clsPool)

    Set moFeet(Index) = oNewValue

End Property

Public Property Get Dependency() As clsPalFromPal

    Set Dependency = moDependency

End Property

Public Property Set Dependency(oNewValue As clsPalFromPal)

    Set moDependency = oNewValue

End Property

Public Property Get MatchCount() As Integer

    MatchCount = mnMatchCount

End Property

Public Property Let MatchCount(nNewValue As Integer)

    mnMatchCount = nNewValue

End Property
