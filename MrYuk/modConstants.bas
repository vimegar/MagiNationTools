Attribute VB_Name = "modConstants"
Option Explicit

Public Enum GB_FILETYPES
    GB_BITMAP = 0
    GB_MAP = 1
    GB_PALETTE = 2
    GB_PATTERN = 3
    GB_VRAM = 4
    GB_BG = 5
    GB_COLLISIONCODES = 6
    GB_COLLISIONMAP = 7
    GB_SPRITEGROUP = 8
End Enum

Public Enum GB_TOOLS
    GB_POINTER = 0
    GB_BRUSH = 1
    GB_MARQUEE = 2
    GB_ZOOM = 3
    GB_SETTER = 4
    GB_BUCKET = 5
    GB_REPLACE = 6
End Enum

Public Enum GB_UPDATETYPES
    GB_ACTIVEEDITOR = 0
    GB_RESOURCECHANGED = 1
End Enum

Public Enum GB_BACKGROUNDTYPES
    GB_PATTERNBG = 1
    GB_RAWBG = 2
End Enum

Public Enum TILETYPES
    tVRAM = 0
    tPat = 1
End Enum

