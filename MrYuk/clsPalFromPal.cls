VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPalFromPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnPoolIndex(4) As Integer

Public Property Let PoolIndex(Index As Integer, nNewValue As Integer)

    mnPoolIndex(Index) = nNewValue

End Property

Public Property Get PoolIndex(Index As Integer) As Integer

    PoolIndex = mnPoolIndex(Index)

End Property
