VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moFeetInMe() As clsPalFromTile
Private moPoolRGB As New clsRGB

Public Property Set FeetInMe(Index As Integer, oNewValue As clsPalFromTile)

    Set moFeetInMe(Index) = oNewValue

End Property

Public Property Get FeetInMe(Index As Integer) As clsPalFromTile

    Set FeetInMe = moFeetInMe(Index)

End Property

Public Property Let FeetInMeCount(nNewValue As Integer)

    ReDim Preserve moFeetInMe(nNewValue)
    
End Property

Public Property Get FeetInMeCount() As Integer

    FeetInMeCount = UBound(moFeetInMe)
    
End Property

Public Property Set PoolRGB(oNewValue As clsRGB)

    Set moPoolRGB = oNewValue

End Property

Public Property Get PoolRGB() As clsRGB

    Set PoolRGB = moPoolRGB

End Property



Private Sub Class_Initialize()

    ReDim moFeetInMe(0)

End Sub


