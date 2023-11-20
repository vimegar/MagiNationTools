Attribute VB_Name = "modPlayer"
Option Explicit

Public Const NUM_WALK_FRAMES = 24
Public Const FRAMES_PER_DIR = 3
Public Const FRAME_SPEED = 0.06

Public Const PLAYER_WALK_SPEED = 1
Public Const PLAYER_RUN_SPEED = 2

Public Const WALK_FRAME_TILE_WIDTH = 2
Public Const WALK_FRAME_TILE_HEIGHT = 3

Public Enum DIR_CONST
    DIR_LEFT = 0
    DIR_UP = 1
    DIR_RIGHT = 2
    DIR_DOWN = 3
End Enum

Public Enum SPEED_CONST
    WALK = 0
    RUN = 1
End Enum

Public Type tPlayer
    xTile As Integer
    yTile As Integer
    Frame As Integer
    FrameList As String
    FrameCursor As Single
    Speed As SPEED_CONST
    Dir As DIR_CONST
End Type

Public gTony As tPlayer

Public gWalkBitmaps() As Object
Public gMaskBitmaps() As Object

Private mnWalkFrameCount As Integer
Public Sub AddWalkFrame(Filename As String)
    
    mnWalkFrameCount = mnWalkFrameCount + 1

    ReDim Preserve gWalkBitmaps(mnWalkFrameCount - 1)
    ReDim Preserve gMaskBitmaps(mnWalkFrameCount - 1)
    
    On Error Resume Next
    Load frmMain.picWalkBitmaps(mnWalkFrameCount - 1)
    Load frmMain.picMaskBitmaps(mnWalkFrameCount - 1)
    On Error GoTo 0
    Set gWalkBitmaps(mnWalkFrameCount - 1) = frmMain.picWalkBitmaps(mnWalkFrameCount - 1)
    Set gMaskBitmaps(mnWalkFrameCount - 1) = frmMain.picMaskBitmaps(mnWalkFrameCount - 1)

    gWalkBitmaps(mnWalkFrameCount - 1).Picture = LoadPicture(Filename)
    gMaskBitmaps(mnWalkFrameCount - 1).Width = gWalkBitmaps(mnWalkFrameCount - 1).Width
    gMaskBitmaps(mnWalkFrameCount - 1).Height = gWalkBitmaps(mnWalkFrameCount - 1).Height
    
    CreateMask gWalkBitmaps(mnWalkFrameCount - 1).hDC, gWalkBitmaps(mnWalkFrameCount - 1).Width, gWalkBitmaps(mnWalkFrameCount - 1).Height, gMaskBitmaps(mnWalkFrameCount - 1).hDC

End Sub

Public Sub DrawPlayer(DestDC As Long)

    On Error GoTo HandleErrors

    BitBlt DestDC, Int(gTony.xPixel), Int(gTony.yPixel) + 8, (WALK_FRAME_TILE_WIDTH * TILE_PIXEL_WIDTH), (WALK_FRAME_TILE_HEIGHT * TILE_PIXEL_HEIGHT), gMaskBitmaps(((gTony.Dir * FRAMES_PER_DIR) + gTony.Frame) + (gTony.Speed * (NUM_WALK_FRAMES / 2))).hDC, 0, 0, SRCAND
    BitBlt DestDC, Int(gTony.xPixel), Int(gTony.yPixel) + 8, (WALK_FRAME_TILE_WIDTH * TILE_PIXEL_WIDTH), (WALK_FRAME_TILE_HEIGHT * TILE_PIXEL_HEIGHT), gWalkBitmaps(((gTony.Dir * FRAMES_PER_DIR) + gTony.Frame) + (gTony.Speed * (NUM_WALK_FRAMES / 2))).hDC, 0, 0, SRCPAINT

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modPlayer:DrawPlayer Error"
End Sub


Public Sub InitPlayer()

    LoadWalkFrames
    
    gTony.FrameList = "0121"
    gTony.xPixel = (4 * TILE_PIXEL_WIDTH)
    gTony.yPixel = (3 * TILE_PIXEL_HEIGHT)

End Sub


Public Sub LoadWalkFrames()

    Dim i As Integer
    
    For i = 0 To (NUM_WALK_FRAMES - 1)
        AddWalkFrame App.Path & "\Graphics\walk" & Format$(CStr(i), "00") & ".bmp"
    Next i

End Sub

