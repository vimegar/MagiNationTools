VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlRed As Long
Private mlGreen As Long
Private mlBlue As Long

Public Property Get Red() As Long

    Red = mlRed

End Property

Public Property Let Red(lNewValue As Long)

    mlRed = lNewValue

End Property

Public Property Get Green() As Long

    Green = mlGreen

End Property

Public Property Let Green(lNewValue As Long)

    mlGreen = lNewValue

End Property

Public Property Get Blue() As Long

    Blue = mlBlue

End Property

Public Property Let Blue(lNewValue As Long)

    mlBlue = lNewValue

End Property


Private Sub Class_Initialize()

    mlRed = 0
    mlGreen = 0
    mlBlue = 0

End Sub


