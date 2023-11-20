VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHotspotRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnType As Integer
Private mnLeft As Integer
Private mnRight As Integer
Private mnTop As Integer
Private mnBottom As Integer

Public Property Get nType() As Integer

    nType = mnType

End Property

Public Property Let nType(nNewValue As Integer)

    mnType = nNewValue

End Property

Public Property Get nLeft() As Integer

    nLeft = mnLeft

End Property

Public Property Let nLeft(nNewValue As Integer)

    mnLeft = nNewValue

End Property

Public Property Get nRight() As Integer

    nRight = mnRight

End Property

Public Property Let nRight(nNewValue As Integer)

    mnRight = nNewValue

End Property

Public Property Get nTop() As Integer

    nTop = mnTop

End Property

Public Property Let nTop(nNewValue As Integer)

    mnTop = nNewValue

End Property

Public Property Get nBottom() As Integer

    nBottom = mnBottom

End Property

Public Property Let nBottom(nNewValue As Integer)

    mnBottom = nNewValue

End Property
