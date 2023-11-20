Attribute VB_Name = "modCreatePatAndVRAM"
Option Explicit

Public Enum COMPARE_TILE_TYPES
    CTPattern = 0
    ctVRAM = 1
End Enum

Public Sub CreateCTArrayFromBMP(sFilename As String, BMPOffscreen As clsOffscreen, oBit As clsGBBitmap, oCTArray As clsCTArray, Optional Progress As ProgressBar, Optional Status As StatusBar)
    
    'get pal from bmp
    'get bit from bmp
    'call mCreateCTArray
    
    On Error GoTo HandleErrors
    
    Dim iPalMap() As Byte
    
    oBit.GBPalette.GetPalFromBMP BMPOffscreen
    oBit.GetBitFromBMP BMPOffscreen, iPalMap, Progress, Status
    
    mCreateCTArray oBit, iPalMap, oCTArray
    
Exit Sub
    
HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:CreateCTArrayFromBMP Error"
End Sub





Public Sub CreatePattern(oCTPat As clsCTArray, oCTVRAM As clsCTArray, oPat As clsGBBackground)

    'place subtile 0 only from oCTPat into subTileArray
    'place all 4 subtiles from oCTVRAM into subTileArray
    'quicksort subTileArray
    'find spans and dependencies
    'output a pattern using pattern comp tiles only

    On Error GoTo HandleErrors

    Dim subTileArray() As clsCTSubtile
    
    'mDebugPrintCTArray oCTPat, "c:\windows\desktop\pat.txt"
    'mDebugPrintCTArray oCTVRAM, "c:\windows\desktop\vrm.txt"
        
    mCreateSubtileArrayPat oCTPat, oCTVRAM, subTileArray
    QuicksortCTSubtile subTileArray, 0, UBound(subTileArray)
    mSetDependencies subTileArray
    mOutputPattern oCTPat, oPat, False
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:CreatePattern Error"
End Sub

Public Sub CreatePatternAndVRAM(oCTArray As clsCTArray, oBit As clsGBBitmap, oVRAM As clsGBVRAM, oPat As clsGBBackground, oVRAMBit As clsGBBitmap)

    'place all 4 subtiles of each CT into subTileArray
    'quicksort subTileArray
    'find spans and dependencies
    'output a VRAM by using CTs with no dependecy
    'output a Pattern using all tiles in CTArray

    On Error GoTo HandleErrors

    Dim subTileArray() As clsCTSubtile

    mCreateSubtileArrayPatAndVRAM oCTArray, subTileArray
    
''debug print
''***************************************************************************************
'Dim i As Integer
'Dim j As Integer
'Dim nFilenum As Integer
'nFilenum = FreeFile
'Open "c:\windows\desktop\subTileArray.txt" For Output As #nFilenum
'For i = 0 To UBound(subTileArray)
'    For j = 0 To 3
'        Print #nFilenum, subTileArray(i).PixelData(j)
'    Next j
'Next i
'Close #nFilenum
'***************************************************************************************
    
    QuicksortCTSubtile subTileArray, 0, UBound(subTileArray)
    
''debug print
''***************************************************************************************
'nFilenum = FreeFile
'Open "c:\windows\desktop\subTileArray.txt" For Output As #nFilenum
'For i = 0 To UBound(subTileArray)
'    For j = 0 To 3
'        Print #nFilenum, subTileArray(i).PixelData(j) & vbTab & vbTab & subTileArray(i).Parent.ID & vbTab & vbTab & subTileArray(i).XFlip & vbTab & vbTab & subTileArray(i).YFlip
'    Next j
'    Print #nFilenum, ""
'Next i
'Close #nFilenum
''***************************************************************************************
    
    mSetDependencies subTileArray
        
''debug display
'Dim frmBit As New frmEditBitmap
'Set frmBit.GBBitmap = oBit
'frmBit.bChanged = False
'frmBit.Show
    
    mOutputVRAM oCTArray, oBit, oVRAM, oVRAMBit
    mOutputPattern oCTArray, oPat, True

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:CreatePatternAndVRAM Error"
End Sub

Private Sub mCompressCT(oCT As clsCT, oBit As clsGBBitmap, xPixel As Integer, yPixel As Integer)

    On Error GoTo HandleErrors
    
    Dim iPixelData(8, 8) As Byte
    Dim iPixelDataX() As Byte
    Dim iPixelDataY() As Byte
    Dim iPixelDataXY() As Byte
    
    Dim X As Integer
    Dim Y As Integer
    
    For Y = 0 To 7
        For X = 0 To 7
            iPixelData(X, Y) = oBit.PixelData(xPixel + X, yPixel + Y)
        Next X
    Next Y
    
    iPixelDataX = mGetFlipBitmap(iPixelData, True, False)
    iPixelDataY = mGetFlipBitmap(iPixelData, False, True)
    iPixelDataXY = mGetFlipBitmap(iPixelData, True, True)

    Dim subTile As Integer
    Dim longNum As Integer
    Dim nCount As Integer
    Dim xIndex As Integer
    Dim yIndex As Integer
    
    For subTile = 0 To 3
    
        Set oCT.Subtiles(subTile) = New clsCTSubtile
        Set oCT.Subtiles(subTile).Parent = oCT
    
        For longNum = 0 To 3
            
            nCount = 0
            oCT.Subtiles(subTile).PixelData(longNum) = -2147483648#
        
            For yIndex = 0 To 1
                For xIndex = 0 To 7
            
                    If subTile = 0 Then
                        oCT.Subtiles(subTile).PixelData(longNum) = oCT.Subtiles(subTile).PixelData(longNum) + (iPixelData(xIndex, yIndex + (longNum * 2)) * (2 ^ nCount))
                    ElseIf subTile = 1 Then
                        oCT.Subtiles(subTile).PixelData(longNum) = oCT.Subtiles(subTile).PixelData(longNum) + (iPixelDataX(xIndex, yIndex + (longNum * 2)) * (2 ^ nCount))
                    ElseIf subTile = 2 Then
                        oCT.Subtiles(subTile).PixelData(longNum) = oCT.Subtiles(subTile).PixelData(longNum) + (iPixelDataY(xIndex, yIndex + (longNum * 2)) * (2 ^ nCount))
                    ElseIf subTile = 3 Then
                        oCT.Subtiles(subTile).PixelData(longNum) = oCT.Subtiles(subTile).PixelData(longNum) + (iPixelDataXY(xIndex, yIndex + (longNum * 2)) * (2 ^ nCount))
                    End If
                    
                    nCount = nCount + 2
                
                Next xIndex
            Next yIndex
            
        Next longNum
    Next subTile
    
    oCT.Subtiles(0).XFlip = False
    oCT.Subtiles(0).YFlip = False
                    
    oCT.Subtiles(1).XFlip = True
    oCT.Subtiles(1).YFlip = False
                    
    oCT.Subtiles(2).XFlip = False
    oCT.Subtiles(2).YFlip = True
                    
    oCT.Subtiles(3).XFlip = True
    oCT.Subtiles(3).YFlip = True
  
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mCompressCT Error"
End Sub


Public Sub CreateCTArrayFromVRAM(oCTArray As clsCTArray, oVRAM As clsGBVRAM)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim nBank As Integer
    Dim nCount As Integer
        
    oVRAM.EnumBitmapFragments
        
    nCount = 0
        
    For nBank = 0 To 1
        For i = 1 To 384
            If Not oVRAM.BitmapFragments(i, nBank).GBBitmap Is Nothing Then
            
                oCTArray.TileCount = nCount + 1
                Set oCTArray.CompareTiles(nCount) = New clsCT
            
                mCompressCT oCTArray.CompareTiles(nCount), oVRAM.BitmapFragments(i, nBank).GBBitmap, oVRAM.BitmapFragments(i, nBank).X, oVRAM.BitmapFragments(i, nBank).Y
                
                oCTArray.CompareTiles(nCount).VRAMAddress = ((i - 1) * 16) + 32768
                oCTArray.CompareTiles(nCount).VRAMBank = nBank
                oCTArray.CompareTiles(nCount).CTType = ctVRAM
                nCount = nCount + 1
            
            End If
        Next i
    Next nBank

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:CreateCTArrayFromVRAM Error"
End Sub


Private Sub mDebugPrintCTArray(oCTArray As clsCTArray, sFilename As String)

    Dim i As Integer
    Dim span As Integer
    Dim nFilenum As Integer
    Dim str As String
    Dim strAddr As String
    
    nFilenum = FreeFile
    span = 1
    
    Open sFilename For Output As #nFilenum
    
    Print #nFilenum, "ID" & vbTab & vbTab & "PatX" & vbTab & vbTab & "PatY" & vbTab & vbTab & "Addr" & vbTab & vbTab & "Bank" & vbTab & vbTab & "Pal" & vbTab & vbTab & "XFlip" & vbTab & vbTab & "YFlip"
    Print #nFilenum, ""
    
    For i = 0 To oCTArray.TileCount - 1
        
        Print #nFilenum, oCTArray.CompareTiles(i).ID & vbTab & vbTab;
        Print #nFilenum, (oCTArray.CompareTiles(i).PixelX \ 8) + 1 & vbTab & vbTab;
        Print #nFilenum, (oCTArray.CompareTiles(i).PixelY \ 8) + 1 & vbTab & vbTab;
        
        If oCTArray.CompareTiles(i).Dependency Is Nothing Then
            Print #nFilenum, Hex(oCTArray.CompareTiles(i).VRAMAddress) & "H" & vbTab & vbTab;
            Print #nFilenum, oCTArray.CompareTiles(i).VRAMBank & vbTab & vbTab;
        Else
            Print #nFilenum, Hex(oCTArray.CompareTiles(i).Dependency.VRAMAddress) & "H" & vbTab & vbTab;
            Print #nFilenum, oCTArray.CompareTiles(i).Dependency.VRAMBank & vbTab & vbTab;
        End If
            
        Print #nFilenum, oCTArray.CompareTiles(i).PaletteID & vbTab & vbTab;
        Print #nFilenum, oCTArray.CompareTiles(i).Subtiles(0).XFlip & vbTab & vbTab;
        Print #nFilenum, oCTArray.CompareTiles(i).Subtiles(0).YFlip & vbTab & vbTab
        
    Next i
    
    Close #nFilenum

End Sub

Private Function mGetFlipBitmap(iData() As Byte, XFlip As Boolean, YFlip As Boolean)

    Dim X As Integer
    Dim Y As Integer
    Dim xCount As Integer
    Dim yCount As Integer
    Dim startX As Integer
    Dim startY As Integer
    Dim endX As Integer
    Dim endY As Integer
    Dim stepX As Integer
    Dim stepY As Integer
    Dim xTile As Integer
    Dim yTile As Integer
    
    If XFlip Then
        startX = 7
        endX = 0
        stepX = -1
    Else
        startX = 0
        endX = 7
        stepX = 1
    End If
    
    If YFlip Then
        startY = 7
        endY = 0
        stepY = -1
    Else
        startY = 0
        endY = 7
        stepY = 1
    End If
    
    ReDim iDataBuffer(UBound(iData, 1), UBound(iData, 2)) As Byte
    
    For yTile = 0 To UBound(iData, 2) \ 8 - 1
        For xTile = 0 To UBound(iData, 1) \ 8 - 1
            
            xCount = 0
            yCount = 0
            
            For Y = startY To endY Step stepY
                For X = startX To endX Step stepX
                    iDataBuffer(xCount + (xTile * 8), yCount + (yTile * 8)) = iData(X + (xTile * 8), Y + (yTile * 8))
                    xCount = xCount + 1
                Next X
                xCount = 0
                yCount = yCount + 1
            Next Y
            
        Next xTile
    Next yTile

    mGetFlipBitmap = iDataBuffer
    
End Function
Private Sub mCreateCTArray(oBit As clsGBBitmap, iPalMap() As Byte, oCTArray As clsCTArray)

    'go through oBit creating all subtiles
    'go through oCTArray copying palmap

    On Error GoTo HandleErrors

    Dim X As Integer
    Dim Y As Integer
    Dim nCount As Integer
    
    nCount = 0
    
    For Y = 0 To oBit.height - 1 Step 8
        For X = 0 To oBit.width - 1 Step 8
    
            oCTArray.TileCount = nCount + 1
            Set oCTArray.CompareTiles(nCount) = New clsCT
            
            mCompressCT oCTArray.CompareTiles(nCount), oBit, X, Y
            
            oCTArray.CompareTiles(nCount).ID = nCount + 1
            oCTArray.CompareTiles(nCount).PixelX = X
            oCTArray.CompareTiles(nCount).PixelY = Y
            
            If iPalMap(X \ 8, Y \ 8) > 0 Then
                oCTArray.CompareTiles(nCount).PaletteID = iPalMap(X \ 8, Y \ 8) - 1
                nCount = nCount + 1
            End If
    
        Next X
    Next Y

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mCreateCTArray Error"
End Sub

Private Sub mCreateSubtileArrayPat(oCTPat As clsCTArray, oCTVRAM As clsCTArray, oSubtileArray() As clsCTSubtile)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer

    ReDim oSubtileArray(oCTPat.TileCount + ((oCTVRAM.TileCount * 4) - 1))
    
    For i = 0 To oCTPat.TileCount - 1
        Set oSubtileArray(i) = oCTPat.CompareTiles(i).Subtiles(0)
    Next i
    
    nCount = oCTPat.TileCount
    
    For i = oCTPat.TileCount To oCTPat.TileCount + oCTVRAM.TileCount - 1
        For j = 0 To 3
            Set oSubtileArray(nCount) = oCTVRAM.CompareTiles(i - oCTPat.TileCount).Subtiles(j)
            nCount = nCount + 1
        Next j
    Next i

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mCreateSubtileArrayPat"
End Sub


Private Sub mCreateSubtileArrayPatAndVRAM(oCTArray As clsCTArray, oSubtileArray() As clsCTSubtile)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer

    ReDim oSubtileArray((oCTArray.TileCount * 4) - 1)
    
    nCount = 0
    
    For i = 0 To oCTArray.TileCount - 1
        For j = 0 To 3
            Set oSubtileArray(nCount) = oCTArray.CompareTiles(i).Subtiles(j)
            nCount = nCount + 1
        Next j
    Next i
    
''debug print
'Dim span As Integer
'Dim nFilenum As Integer
'Dim str As String
'nFilenum = FreeFile
'Open "c:\windows\desktop\unsorted.txt" For Output As #nFilenum
'Print #nFilenum, "ID" & vbTab & vbTab & "DepID" & vbTab & vbTab & "XFlip" & vbTab & vbTab & "YFlip" & vbTab & vbTab & "Pix1" & vbTab & vbTab & vbTab & "Pix1" & vbTab & vbTab & vbTab & "Pix2" & vbTab & vbTab & vbTab & "Pix3"
'For i = 0 To UBound(oSubtileArray)
'    If oSubtileArray(i).Parent.Dependency Is Nothing Then
'        str = "-1"
'    Else
'        str = CStr(oSubtileArray(i).Parent.Dependency.ID)
'    End If
'    Print #nFilenum, oSubtileArray(i).Parent.ID & vbTab & vbTab & str & vbTab & vbTab & oSubtileArray(i).XFlip & vbTab & vbTab & oSubtileArray(i).YFlip & vbTab & vbTab & oSubtileArray(i).PixelData(0) & vbTab & vbTab & oSubtileArray(i).PixelData(1) & vbTab & vbTab & oSubtileArray(i).PixelData(2) & vbTab & vbTab & oSubtileArray(i).PixelData(3)
'Next i
'Close #nFilenum
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mCreateSubtileArrayPatAndVRAM Error"
End Sub


Private Sub mOutputPattern(oCTArray As clsCTArray, oPat As clsGBBackground, OutputingVRAM As Boolean)

    'output a Pattern using all tiles in CTArray

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim xPat As Integer
    Dim yPat As Integer
    
    For i = 0 To oCTArray.TileCount - 1
            
        If oCTArray.CompareTiles(i).CTType = CTPattern Then
            
            If Not OutputingVRAM Then
                If oCTArray.CompareTiles(i).Dependency Is Nothing Then
                    AddTileGetterError
                    TileGetterErrors(TileGetterErrorCount).nType = NoPixelMatch
                    TileGetterErrors(TileGetterErrorCount).X = oCTArray.CompareTiles(i).PixelX \ 8
                    TileGetterErrors(TileGetterErrorCount).Y = oCTArray.CompareTiles(i).PixelY \ 8
                End If
            End If
            
            xPat = (oCTArray.CompareTiles(i).PixelX \ 8) + 1
            yPat = (oCTArray.CompareTiles(i).PixelY \ 8) + 1
                
            If oCTArray.CompareTiles(i).Dependency Is Nothing Then
                oPat.VRAMEntryAddress(xPat, yPat) = oCTArray.CompareTiles(i).VRAMAddress
                oPat.VRAMEntryBank(xPat, yPat) = oCTArray.CompareTiles(i).VRAMBank
            Else
                oPat.VRAMEntryAddress(xPat, yPat) = oCTArray.CompareTiles(i).Dependency.VRAMAddress
                oPat.VRAMEntryBank(xPat, yPat) = oCTArray.CompareTiles(i).Dependency.VRAMBank
            End If
                
            oPat.PaletteID(xPat, yPat) = oCTArray.CompareTiles(i).PaletteID
            oPat.XFlip(xPat, yPat) = oCTArray.CompareTiles(i).Subtiles(0).XFlip
            oPat.YFlip(xPat, yPat) = oCTArray.CompareTiles(i).Subtiles(0).YFlip
            
        End If

    Next i

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mOutputPattern Error"
End Sub

Private Sub mOutputVRAM(oCTArray As clsCTArray, oBit As clsGBBitmap, oVRAM As clsGBVRAM, oVRAMBit As clsGBBitmap)

    'output a VRAM by using CTs with no dependecy

    Dim i As Integer
    Dim nCount As Integer
    Dim tileWidth As Integer
    Dim tileHeight As Integer
    
    nCount = 0
    
    For i = 0 To oCTArray.TileCount - 1
        If oCTArray.CompareTiles(i).Dependency Is Nothing Then
            nCount = nCount + 1
        End If
    Next i
    
    If nCount > 16 Then
        tileWidth = 16
        tileHeight = (nCount \ 16) + 1
    Else
        tileWidth = nCount
        tileHeight = 1
    End If
        
    oVRAMBit.Offscreen.Create tileWidth * 8, tileHeight * 8
    oVRAMBit.ResizePixelData tileWidth * 8, tileHeight * 8
    
    Dim xDest As Integer
    Dim yDest As Integer
    Dim xPixel As Integer
    Dim yPixel As Integer
    Dim vramCount As Long
    
    vramCount = 36864
    nCount = 0
    
    For i = 0 To oCTArray.TileCount - 1
        
        If oCTArray.CompareTiles(i).Dependency Is Nothing Then
        
            xDest = (nCount Mod 16) * 8
            yDest = (nCount \ 16) * 8
            nCount = nCount + 1
        
            For yPixel = 0 To 7
                For xPixel = 0 To 7
                    oVRAMBit.PixelData(xDest + xPixel, yDest + yPixel) = oBit.PixelData(oCTArray.CompareTiles(i).PixelX + xPixel, oCTArray.CompareTiles(i).PixelY + yPixel)
                Next xPixel
            Next yPixel
            
            oCTArray.CompareTiles(i).VRAMAddress = vramCount
            vramCount = vramCount + 16
        
        End If
                    
    Next i
    
    oVRAMBit.TileCount = nCount 'oCTArray.TileCount
    oVRAMBit.RenderPixels
    
End Sub

Private Sub mSetDependencies(oSubtileArray() As clsCTSubtile)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim SpanStarts() As Integer
    Dim SpanCount As Integer
    
    ReDim SpanStarts(1)
    SpanStarts(0) = 0
    SpanCount = 1
    
    For i = 0 To UBound(oSubtileArray) - 1
    
        If oSubtileArray(i).PixelData(0) = oSubtileArray(i + 1).PixelData(0) Then
            If oSubtileArray(i).PixelData(1) = oSubtileArray(i + 1).PixelData(1) Then
                If oSubtileArray(i).PixelData(2) = oSubtileArray(i + 1).PixelData(2) Then
                    If oSubtileArray(i).PixelData(3) = oSubtileArray(i + 1).PixelData(3) Then
                        GoTo NextSubTile
                    End If
                End If
            End If
        End If
        
        SpanCount = SpanCount + 1
        ReDim Preserve SpanStarts(SpanCount)
        SpanStarts(SpanCount - 1) = i + 1
    
NextSubTile:
        
    Next i

    Dim SpanNum As Integer
    Dim SpanCursor As Integer
    Dim SpanSubtile As clsCTSubtile
    Dim SpanEnd As Integer
    
    For SpanNum = 0 To SpanCount - 1
        If SpanNum = SpanCount - 1 Then
            SpanEnd = UBound(oSubtileArray)
        Else
            SpanEnd = SpanStarts(SpanNum + 1) - 1
        End If
        
        For SpanCursor = SpanStarts(SpanNum) To SpanEnd
            If oSubtileArray(SpanCursor).Parent.Dependency Is Nothing Then
                Set SpanSubtile = oSubtileArray(SpanCursor)
                Exit For
            End If
        Next SpanCursor
        
        For SpanCursor = SpanStarts(SpanNum) To SpanEnd
            If oSubtileArray(SpanCursor).Parent.CTType = ctVRAM Then
                Set SpanSubtile = oSubtileArray(SpanCursor)
                Exit For
            End If
        Next SpanCursor
        
        If SpanSubtile Is Nothing Then
            GoTo NextSpan
        End If
        
        For SpanCursor = SpanStarts(SpanNum) To SpanEnd
            If Not oSubtileArray(SpanCursor).Parent Is SpanSubtile.Parent Then
                If oSubtileArray(SpanCursor).Parent.CTType = CTPattern Then
                    Set oSubtileArray(SpanCursor).Parent.Dependency = SpanSubtile.Parent
                    oSubtileArray(SpanCursor).XFlip = SpanSubtile.XFlip
                    oSubtileArray(SpanCursor).YFlip = SpanSubtile.YFlip
                End If
            End If
        Next SpanCursor
        
NextSpan:
    Next SpanNum
    
''debug print
'Dim span As Integer
'Dim nFilenum As Integer
'Dim str As String
'nFilenum = FreeFile
'span = 1
'Open "c:\windows\desktop\dependencies.txt" For Output As #nFilenum
'Print #nFilenum, "Type" & vbTab & vbTab & "ID" & vbTab & vbTab & "DepID" & vbTab & vbTab & "XFlip" & vbTab & vbTab & "YFlip" & vbTab & vbTab & "Pix1" & vbTab & vbTab & vbTab & "Pix1" & vbTab & vbTab & vbTab & "Pix2" & vbTab & vbTab & vbTab & "Pix3"
'For i = 0 To UBound(oSubtileArray)
'    If oSubtileArray(i).Parent.Dependency Is Nothing Then
'        str = "-1"
'    Else
'        str = CStr(oSubtileArray(i).Parent.Dependency.ID)
'    End If
'    Print #nFilenum, Mid$("Pat VRAM", (oSubtileArray(i).Parent.CTType * 4) + 1, 4) & vbTab & vbTab & oSubtileArray(i).Parent.ID & vbTab & vbTab & str & vbTab & vbTab & oSubtileArray(i).XFlip & vbTab & vbTab & oSubtileArray(i).YFlip & vbTab & vbTab & oSubtileArray(i).PixelData(0) & vbTab & vbTab & oSubtileArray(i).PixelData(1) & vbTab & vbTab & oSubtileArray(i).PixelData(2) & vbTab & vbTab & oSubtileArray(i).PixelData(3)
'    If i = SpanStarts(span) - 1 Then
'        span = span + 1
'        Print #nFilenum, ""
'    End If
'Next i
'Close #nFilenum
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modCreatePatAndVRAM:mSetDependencies Error"
End Sub


