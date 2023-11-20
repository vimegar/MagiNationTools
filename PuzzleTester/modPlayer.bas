Attribute VB_Name = "modPlayer"
Option Explicit

Public Const NUM_PLAYER_FRAMES = 4

Public Const PLAYER_WIDTH_IN_TILES = 2
Public Const PLAYER_HEIGHT_IN_TILES = 3

Public Const PLAYER_XOFFSET_IN_PIXELS = 0
Public Const PLAYER_YOFFSET_IN_PIXELS = 8

Public Enum DIRECTION_CONSTANTS
    DIR_LEFT = 0
    DIR_UP = 1
    DIR_RIGHT = 2
    DIR_DOWN = 3
End Enum

Public gPlayer As clsPlayer

Public gWalkBitmaps(NUM_PLAYER_FRAMES - 1) As Object
Public gMaskBitmaps(NUM_PLAYER_FRAMES - 1) As Object
Public Sub DrawPlayer(BufferDC As Long)

    BitBlt BufferDC, (((DISPLAY_WIDTH_IN_TILES \ 2) - 1) * TILE_WIDTH_IN_PIXELS) + PLAYER_XOFFSET_IN_PIXELS, (((DISPLAY_HEIGHT_IN_TILES \ 2) - 1) * TILE_HEIGHT_IN_PIXELS) + PLAYER_YOFFSET_IN_PIXELS, PLAYER_WIDTH_IN_TILES * TILE_WIDTH_IN_PIXELS, PLAYER_HEIGHT_IN_TILES * TILE_HEIGHT_IN_PIXELS, gMaskBitmaps(gPlayer.Direction).hdc, 0, 0, vbSrcAnd
    BitBlt BufferDC, (((DISPLAY_WIDTH_IN_TILES \ 2) - 1) * TILE_WIDTH_IN_PIXELS) + PLAYER_XOFFSET_IN_PIXELS, (((DISPLAY_HEIGHT_IN_TILES \ 2) - 1) * TILE_HEIGHT_IN_PIXELS) + PLAYER_YOFFSET_IN_PIXELS, PLAYER_WIDTH_IN_TILES * TILE_WIDTH_IN_PIXELS, PLAYER_HEIGHT_IN_TILES * TILE_HEIGHT_IN_PIXELS, gWalkBitmaps(gPlayer.Direction).hdc, 0, 0, vbSrcPaint
    
End Sub

Public Sub CreateMask(SrcDC As Long, SrcPixelWidth As Integer, SrcPixelHeight As Integer, MaskDC As Long)

    Dim x As Integer
    Dim y As Integer

    For y = 0 To (SrcPixelHeight - 1)
        For x = 0 To (SrcPixelWidth - 1)
            If GetPixel(SrcDC, x, y) = vbBlack Then
                SetPixel MaskDC, x, y, vbWhite
            Else
                SetPixel MaskDC, x, y, vbBlack
            End If
        Next x
    Next y

End Sub

Public Sub InitPlayer()

    Set gPlayer = New clsPlayer
    
    gPlayer.XInTiles = 5
    gPlayer.YInTiles = 5
    
    LoadPlayerGraphics
    
End Sub

Public Sub LoadPlayerGraphics()

    Dim i As Integer
    
    For i = 0 To (NUM_PLAYER_FRAMES - 1)
        
        If i > 0 Then
            Load frmMain.picPlayer(i)
            Load frmMain.picPlayerMask(i)
        End If
        
        Set gWalkBitmaps(i) = frmMain.picPlayer(i)
        Set gMaskBitmaps(i) = frmMain.picPlayerMask(i)
        
        gWalkBitmaps(i).Picture = LoadPicture(App.Path & "\Graphics\walk" & Format$(CStr(i), "00") & ".bmp")
        CreateMask gWalkBitmaps(i).hdc, gWalkBitmaps(i).Width, gWalkBitmaps(i).Height, gMaskBitmaps(i).hdc
        
    Next i

End Sub


