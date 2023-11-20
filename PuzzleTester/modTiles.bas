Attribute VB_Name = "modTiles"
Option Explicit

Public Const NUM_TILES = 15

Public gTileBitmaps(NUM_TILES - 1) As Object
Public Sub InitTiles()

    Dim i As Integer
    
    For i = 0 To (NUM_TILES - 1)
        
        If i > 0 Then
            Load frmMain.picTiles(i)
        End If
        
        Set gTileBitmaps(i) = frmMain.picTiles(i)
        gTileBitmaps(i).Picture = LoadPicture(App.Path & "\Graphics\tile" & Format$(CStr(i), "00") & ".bmp")
        
    Next i

End Sub


