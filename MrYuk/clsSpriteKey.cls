VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpriteKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msprite As New clsGBSpriteGroup
Private mnFrameCount As Integer
Private msSpriteFile As String

Public Property Get Sprite() As clsSpriteTileGroup

    Set Sprite = msprite.Sprites(1)

End Property

Public Property Set Sprite(oNewValue As clsSpriteTileGroup)

    Set msprite.Sprites(1) = oNewValue

End Property

Public Property Get FrameCount() As Integer

    FrameCount = mnFrameCount

End Property

Public Property Let FrameCount(nNewValue As Integer)

    mnFrameCount = nNewValue

End Property

Public Property Let sSpriteFile(sNewValue As String)

    msSpriteFile = sNewValue

End Property

Public Property Get sSpriteFile() As String

    sSpriteFile = msSpriteFile

End Property

Private Sub Class_Initialize()

    msprite.AddSprite

End Sub


