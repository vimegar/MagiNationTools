Attribute VB_Name = "modTiles"
Option Explicit

Public Const NUM_TILES = 15

Public gTileBitmaps() As Object

Private mnTileCount As Integer

Public Sub AddTile(Filename As String)

    mnTileCount = mnTileCount + 1
    
    ReDim Preserve gTileBitmaps(mnTileCount - 1)
    
    On Error Resume Next
    Load frmMain.picTileBitmaps(mnTileCount - 1)
    Set gTileBitmaps(mnTileCount - 1) = frmMain.picTileBitmaps(mnTileCount - 1)
    On Error GoTo 0
    
    gTileBitmaps(mnTileCount - 1).Picture = LoadPicture(Filename)

End Sub


Public Sub InitTiles()

    Dim i As Integer
    
    For i = 0 To (NUM_TILES - 1)
        AddTile App.Path & "\Graphics\tile" & Format$(CStr(i), "00") & ".bmp"
    Next i

End Sub


