VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBitmapFragment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnX As Integer
Private mnY As Integer
Private mGBBitmap As clsGBBitmap

Public Sub Delete()

    If Not mGBBitmap Is Nothing Then
        mGBBitmap.Delete
    End If

End Sub


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

Public Property Get GBBitmap() As clsGBBitmap

    Set GBBitmap = mGBBitmap

End Property

Public Property Set GBBitmap(oNewValue As clsGBBitmap)

    Set mGBBitmap = oNewValue

End Property

