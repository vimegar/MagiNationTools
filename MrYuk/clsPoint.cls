VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlX As Long
Private mlY As Long

Public Property Get x() As Long

    x = mlX

End Property

Public Property Let x(lNewValue As Long)

    mlX = lNewValue

End Property

Public Property Get y() As Long

    y = mlY

End Property

Public Property Let y(lNewValue As Long)

    mlY = lNewValue

End Property

