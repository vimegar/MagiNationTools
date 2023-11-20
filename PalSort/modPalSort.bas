Attribute VB_Name = "modPalSort"
Option Explicit

Public Sub GetPalMapFromPools(oPools As clsPool, oPalsFromPal() As clsPalFromPal, oPalsFromTile() As clsPalFromTile, iPalMap() As Byte)

End Sub

Public Sub GetPalsFromTiles(BMPOffscreen As clsOffscreen, oPalsFromTile() As clsPalFromTile)

    Dim xTile As Integer
    Dim yTile As Integer
    Dim xPixel As Integer
    Dim yPixel As Integer
    Dim lPixelVal As Long
    
    For yTile = 0 To (BMPOffscreen.Height \ 8) - 1
        For xTile = 0 To (BMPOffscreen.Width \ 8) - 1
        
            For yPixel = 0 To 7
                For xPixel = 0 To 7
                    
                    lPixelVal = BMPOffscreen.GetPixel((xTile * 8) + xPixel, (yTile * 8) + yPixel)
                    
                Next xPixel
            Next yPixel
        
        Next xTile
    Next yTile

End Sub


Public Sub GetPoolsFromPal(oPal As clsGBPalette, oPools() As clsPool, oPalsFromPal() As clsPalFromPal)

    'create pool array
    'set up indexes in oPalsFromPal
    
    Dim i As Integer
    
    For i = 0 To 31
        
    Next i

End Sub

Public Function GetRGBFromLong(lColor As Long) As clsRGB

'***************************************************************************
'   Return a clsRGB from a long color value
'***************************************************************************

    Dim dummy As Long
    
    Set GetRGBFromLong = New clsRGB
    
    GetRGBFromLong.Red = lColor And &HFF
    GetRGBFromLong.Green = (Fix(lColor / 256)) And &HFF
    GetRGBFromLong.Blue = (Fix(lColor / 65536)) And &HFF
    
End Function
