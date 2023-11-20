VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpriteTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnXOffset As Integer      '-128 to 127
Private mnYOffset As Integer      '-128 to 127
Private miTileID As Byte          '0 to 255
Private miBank As Byte            '0 to 1
Private miXFlip As Byte           '0 to 1
Private miYFlip As Byte           '0 to 1
Private miPalID As Byte           '0 to 7
Private miPriority As Byte        '0 to 1
Private mnBitmapFragmentIndex As Integer

Public Property Let BitmapFragmentIndex(nNewValue As Integer)

    mnBitmapFragmentIndex = nNewValue

End Property

Public Property Get BitmapFragmentIndex() As Integer
    
    BitmapFragmentIndex = mnBitmapFragmentIndex
    
End Property

Public Property Let PalID(iNewValue As Byte)

    miPalID = iNewValue

End Property

Public Property Let Bank(iNewValue As Byte)

    miBank = iNewValue

End Property

Public Property Let Priority(iNewValue As Byte)

    miPriority = iNewValue

End Property

Public Property Get Priority() As Byte

    Priority = miPriority

End Property

Public Property Let TileID(iNewValue As Byte)

    miTileID = iNewValue

End Property

Public Property Let XFlip(iNewValue As Byte)

    miXFlip = iNewValue

End Property

Public Property Let YFlip(iNewValue As Byte)

    miYFlip = iNewValue

End Property

Public Property Let XOffset(nNewValue As Integer)

    mnXOffset = nNewValue

End Property

Public Property Let yOffset(nNewValue As Integer)

    mnYOffset = nNewValue

End Property

Public Property Get TileID() As Byte

    TileID = miTileID

End Property

Public Property Get Bank() As Byte

    Bank = miBank

End Property

Public Property Get PalID() As Byte

    PalID = miPalID

End Property

Public Property Get XFlip() As Byte

    XFlip = miXFlip

End Property

Public Property Get YFlip() As Byte

    YFlip = miYFlip

End Property

Public Property Get XOffset() As Integer

    XOffset = mnXOffset

End Property

Public Property Get yOffset() As Integer

    yOffset = mnYOffset

End Property

