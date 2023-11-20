VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbWalkable As Boolean

Public Property Let Walkable(bNewValue As Boolean)

    mbWalkable = bNewValue

End Property

Public Property Get Walkable() As Boolean

    Walkable = mbWalkable

End Property

