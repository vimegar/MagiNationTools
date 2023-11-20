Attribute VB_Name = "modPalSort"
Option Explicit

Public Sub GetPalMapFromPools(oPools() As clsPool, oPalsFromPal() As clsPalFromPal, oPalsFromTile() As clsPalFromTile, iPalMap() As Byte, Optional Progress As ProgressBar)

    'for each PalsFromPal
        'for each poolIndex in each PalsFromPal
            'increment the MatchCount property of PalsFromTile with their feet in pool
        'next
        'for each PalsFromTile
            'if MatchCount = -1 then exit loop
            'if MatchCount = ColorCount then
                'set iPalMap(PalsFromTile.SourceX, PalsFromTile.SourceY) = current pal index
                
                'For Each color In PalFromTile
                    'Visit the pool containing the color
                        'Visit the index of the color and set the color to nothing
            'else
                'set MatchCount = 0
            'end if
        'next
    'next

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim j As Integer
    Dim pfpCursor As Integer
    Dim pftCursor As Integer
    
    For pfpCursor = 0 To UBound(oPalsFromPal) - 1
        
        For i = 0 To 3
            For j = 0 To oPools(oPalsFromPal(pfpCursor).PoolIndex(i)).FeetInMeCount - 1
                If Not oPools(oPalsFromPal(pfpCursor).PoolIndex(i)).FeetInMe(j) Is Nothing Then
                    oPools(oPalsFromPal(pfpCursor).PoolIndex(i)).FeetInMe(j).MatchCount = oPools(oPalsFromPal(pfpCursor).PoolIndex(i)).FeetInMe(j).MatchCount + 1
                End If
            Next j
        Next i
        
        For pftCursor = 0 To UBound(oPalsFromTile)
            If oPalsFromTile(pftCursor) Is Nothing Then
                GoTo NextPFT
            End If
            If oPalsFromTile(pftCursor).MatchCount = -1 Then
                GoTo NextPFT
            End If
            If oPalsFromTile(pftCursor).MatchCount = oPalsFromTile(pftCursor).ColorCount Then
                iPalMap(oPalsFromTile(pftCursor).SourceX, oPalsFromTile(pftCursor).SourceY) = pfpCursor + 1
                oPalsFromTile(pftCursor).MatchCount = -1
                
                'set matched parent's legs to nothing
                For i = 0 To oPalsFromTile(pftCursor).ColorCount - 1
                    Set oPalsFromTile(pftCursor).Feet(i).FeetInMe(oPalsFromTile(pftCursor).FeetIndex(i)) = Nothing
                Next i
            Else
                oPalsFromTile(pftCursor).MatchCount = 0
            End If
NextPFT:
        Next pftCursor
        
        If Not Progress Is Nothing Then
            Progress.value = Progress.value + 1
        End If
    Next pfpCursor

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modPalSort:GetPalMapFromPools Error"
End Sub

Public Sub GetPalsFromTiles(BMPOffscreen As clsOffscreen, oPools() As clsPool, oPalsFromTile() As clsPalFromTile, Optional Progress As ProgressBar)

    On Error GoTo HandleErrors

    Dim xTile As Integer
    Dim yTile As Integer
    Dim xPixel As Integer
    Dim yPixel As Integer
    Dim pixelRGB As New clsRGB
    Dim matchedColors() As New clsRGB
    Dim ColorCount As Integer
    Dim colorExists As Boolean
    Dim TileCount As Integer
    Dim matchedPool As clsPool
    Dim i As Integer
    Dim j As Integer
    
    For yTile = 0 To (BMPOffscreen.height \ 8) - 1
        For xTile = 0 To (BMPOffscreen.width \ 8) - 1
        
            ColorCount = 0
        
            For yPixel = 0 To 7
                For xPixel = 0 To 7
                    
                    Set pixelRGB = GetRGBFromLong(BMPOffscreen.GetPixel((xTile * 8) + xPixel, (yTile * 8) + yPixel))
                    
                    colorExists = False
                    
                    For i = 0 To ColorCount - 1
                        If pixelRGB.Red = matchedColors(i).Red And pixelRGB.Green = matchedColors(i).Green And pixelRGB.Blue = matchedColors(i).Blue Then
                            colorExists = True
                            Exit For
                        End If
                    Next i
                    
                    If Not colorExists Then
                        ColorCount = ColorCount + 1
                        ReDim Preserve matchedColors(ColorCount - 1)
                        matchedColors(ColorCount - 1).Red = pixelRGB.Red
                        matchedColors(ColorCount - 1).Green = pixelRGB.Green
                        matchedColors(ColorCount - 1).Blue = pixelRGB.Blue
                    End If
                    
                Next xPixel
            Next yPixel
            
            TileCount = TileCount + 1
            ReDim Preserve oPalsFromTile(TileCount - 1)
            
            If ColorCount > 4 Then
                'MsgBox "There is an error in the source bitmap.  The tile located at " & CStr(xTile + 1) & ", " & CStr(yTile + 1) & " contains " & CStr(ColorCount) & " colors.  You must reduce the amount of colors to 4 per tile!", vbCritical, "Palette Error"
                AddTileGetterError
                TileGetterErrors(TileGetterErrorCount).ColorCount = ColorCount
                TileGetterErrors(TileGetterErrorCount).nType = TooManyColors
                TileGetterErrors(TileGetterErrorCount).X = xTile
                TileGetterErrors(TileGetterErrorCount).Y = yTile
                GoTo NextTile
            End If
            
            For i = 0 To ColorCount - 1
                For j = 0 To UBound(oPools) - 1
                    If (matchedColors(i).Red \ 8) = oPools(j).PoolRGB.Red And (matchedColors(i).Green \ 8) = oPools(j).PoolRGB.Green And (matchedColors(i).Blue \ 8) = oPools(j).PoolRGB.Blue Then
                        Set matchedPool = oPools(j)
                        Exit For
                    End If
                Next j
                If oPalsFromTile(TileCount - 1) Is Nothing Then
                    Set oPalsFromTile(TileCount - 1) = New clsPalFromTile
                End If
                If matchedPool Is Nothing Then
                    GoTo NextTile
                End If
                Set oPalsFromTile(TileCount - 1).Feet(i) = matchedPool
                matchedPool.FeetInMeCount = matchedPool.FeetInMeCount + 1
                Set matchedPool.FeetInMe(matchedPool.FeetInMeCount - 1) = oPalsFromTile(TileCount - 1)
                oPalsFromTile(TileCount - 1).FeetIndex(i) = matchedPool.FeetInMeCount - 1
                oPalsFromTile(TileCount - 1).SourceX = xTile
                oPalsFromTile(TileCount - 1).SourceY = yTile
                oPalsFromTile(TileCount - 1).ColorCount = ColorCount
            Next i
NextTile:
        Next xTile
        If Not Progress Is Nothing Then
            Progress.value = Progress.value + 1
        End If
    
    Next yTile
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modPalSort:GetPalsFromTiles Error"
End Sub


Public Sub GetPoolsFromPal(oPal As clsGBPalette, oPools() As clsPool, oPalsFromPal() As clsPalFromPal, Optional Progress As ProgressBar)

    'create pool array
    'set up indexes in oPalsFromPal
    
    On Error GoTo HandleErrors
    
    Dim i As Integer
    Dim colorIsOld As Boolean
    Dim colorCursor As Integer
    Dim poolCursor As Integer
    Dim poolCount As Integer
    
    ReDim oPalsFromPal(8)
    For i = 0 To 7
        Set oPalsFromPal(i) = New clsPalFromPal
    Next i
    
    poolCount = 0
    ReDim oPools(0)
    Set oPools(0) = New clsPool
    
    oPools(0).PoolRGB.Red = -1
    oPools(0).PoolRGB.Green = -1
    oPools(0).PoolRGB.Blue = -1
    
    For colorCursor = 1 To 32
        
        'check for existing pool
        colorIsOld = False
        For poolCursor = 0 To (poolCount - 1)
            If oPal.Colors(colorCursor).Red = oPools(poolCursor).PoolRGB.Red And oPal.Colors(colorCursor).Green = oPools(poolCursor).PoolRGB.Green And oPal.Colors(colorCursor).Blue = oPools(poolCursor).PoolRGB.Blue Then
                oPalsFromPal((colorCursor - 1) \ 4).PoolIndex((colorCursor - 1) Mod 4) = poolCursor
                colorIsOld = True
                Exit For
            End If
        Next poolCursor
        
        'if color is new then add new pool
        If Not colorIsOld Then
            poolCount = poolCount + 1
            ReDim Preserve oPools(poolCount - 1)
            Set oPools(poolCount - 1) = New clsPool
            oPools(poolCount - 1).PoolRGB.Red = oPal.Colors(colorCursor).Red
            oPools(poolCount - 1).PoolRGB.Green = oPal.Colors(colorCursor).Green
            oPools(poolCount - 1).PoolRGB.Blue = oPal.Colors(colorCursor).Blue
            oPalsFromPal((colorCursor - 1) \ 4).PoolIndex((colorCursor - 1) Mod 4) = poolCount - 1
        End If
        
        If Not Progress Is Nothing Then
            Progress.value = Progress.value + 1
        End If
        
    Next colorCursor

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modPalSort:GetPoolsFromPal Error"
End Sub

