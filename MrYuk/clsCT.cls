VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnPixelX As Integer
Private mnPixelY As Integer

Private mnID As Integer
Private miPaletteID As Byte
Private mlVRAMAddress As Long
Private miVRAMBank As Byte

Private moDependency As clsCT
Private mnType As COMPARE_TILE_TYPES

Private moSubtiles(4) As clsCTSubtile

Public Property Set Dependency(oNewValue As clsCT)

    Set moDependency = oNewValue

End Property

Public Property Get Dependency() As clsCT

    Set Dependency = moDependency

End Property



Public Property Let ID(nNewValue As Integer)

    mnID = nNewValue

End Property

Public Property Get ID() As Integer

    ID = mnID

End Property

Public Property Set Subtiles(Index As Integer, oNewValue As clsCTSubtile)

    Set moSubtiles(Index) = oNewValue

End Property

Public Property Get Subtiles(Index As Integer) As clsCTSubtile

    Set Subtiles = moSubtiles(Index)

End Property

Public Property Let CTType(tNewValue As COMPARE_TILE_TYPES)

    mnType = tNewValue

End Property

Public Property Get CTType() As COMPARE_TILE_TYPES

     CTType = mnType

End Property


Public Property Let VRAMBank(iNewValue As Byte)
    
    miVRAMBank = iNewValue

End Property

Public Property Get VRAMBank() As Byte

    VRAMBank = miVRAMBank

End Property

Public Property Get VRAMAddress() As Long

    VRAMAddress = mlVRAMAddress

End Property

Public Property Let VRAMAddress(lNewValue As Long)

    mlVRAMAddress = lNewValue

End Property

Public Property Let PaletteID(iNewValue As Byte)

    miPaletteID = iNewValue

End Property

Public Property Get PaletteID() As Byte

    PaletteID = miPaletteID

End Property

Public Property Get PixelX() As Integer

    PixelX = mnPixelX

End Property

Public Property Let PixelX(nNewValue As Integer)

    mnPixelX = nNewValue

End Property

Public Property Get PixelY() As Integer

    PixelY = mnPixelY

End Property

Public Property Let PixelY(nNewValue As Integer)

    mnPixelY = nNewValue

End Property

