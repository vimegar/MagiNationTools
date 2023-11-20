VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'   VB Interface Setup
'***************************************************************************

    Option Explicit
    Implements intResource
    Implements intResourceClient

'***************************************************************************
'   Class properties
'***************************************************************************

    Private miPixelData() As Byte
    Private mOffscreen As New clsOffscreen
    Private mGBPalette As New clsGBPalette
    Private mnTileCount As Integer
    Private mbClipping As Boolean
    Private msParentPath As String

'***************************************************************************
'   Resource cache object pointers
'***************************************************************************

    Private mResourceCacheObject As clsResourceCacheObject

'***************************************************************************
'   Internal variables
'***************************************************************************

    Private mlGrayScale(3) As Long
Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property


Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Public Property Let bClipping(bNewValue As Boolean)

    mbClipping = bNewValue

End Property

Public Property Get bClipping() As Boolean

    bClipping = mbClipping

End Property

Public Property Get GBPalette() As clsGBPalette

    Set GBPalette = mGBPalette

End Property

Public Property Set GBPalette(oNewValue As clsGBPalette)

    Set mGBPalette = oNewValue

End Property
Public Sub Delete()

    intResource_Delete

End Sub

Public Sub GetBitFromBMP(BMPOffscreen As clsOffscreen, PalMap() As Byte, Optional Progress As ProgressBar, Optional Status As StatusBar, Optional StatusText As String)

'***************************************************************************
'   Unpack a Windows Bitmap and save it in the GBBitmap class variable
'***************************************************************************

    On Error GoTo HandleErrors

    Dim X As Long
    Dim Y As Long
    Dim flag As Boolean
    Dim nRGB As New clsRGB
    Dim lCropColor As Long
    
    'get crop color
    lCropColor = BMPOffscreen.GetPixel(32, 0)
    
    For X = 0 To BMPOffscreen.width - 1
        BMPOffscreen.SetPixel X, BMPOffscreen.height - 1, lCropColor
    Next X
    
    For Y = 0 To BMPOffscreen.height - 1
        BMPOffscreen.SetPixel BMPOffscreen.width - 1, Y, lCropColor
    Next Y
    
    'get bitmap data
    ReDim miPixelData(BMPOffscreen.width * BMPOffscreen.height)
    
    Dim dWidth As Integer
    Dim dHeight As Integer
    
    dWidth = BMPOffscreen.width
    dHeight = BMPOffscreen.height
    
    If Not Progress Is Nothing Then
        Progress.value = 0
        Progress.Max = BMPOffscreen.height
    End If
    
    If Not Status Is Nothing Then
        If StatusText = "" Then
            Status.Panels(1).Text = "Unpacking bitmap..."
        Else
            Status.Panels(1).Text = StatusText
        End If
    End If
    
    For Y = 8 To BMPOffscreen.height - 1
        If BMPOffscreen.GetPixel(0, Y) = lCropColor Then
            dHeight = Y - 8
            If Not Progress Is Nothing Then
                Do Until Progress.value = Progress.Max
                    Progress.value = Progress.value + 1
                Loop
            End If
            Exit For
        End If
        For X = 0 To BMPOffscreen.width - 1
            If (BMPOffscreen.GetPixel(X, 8) = lCropColor) And flag = False Then
                flag = True
                dWidth = X
                Exit For
            End If
            Set nRGB = GetRGBFromLong(BMPOffscreen.GetPixel(X, Y))
            miPixelData(X + ((Y - 8) * BMPOffscreen.width)) = GetGrayFromRGB(CLng(nRGB.Red), CLng(nRGB.Green), CLng(nRGB.Blue))
            mOffscreen.SetPixel X, (Y - 8), mlGrayScale(miPixelData(X + ((Y - 8) * BMPOffscreen.width)))
            
            'miPixelData(x + ((y - 8) * dWidth)) = GetGrayFromRGB(CLng(nRGB.Red), CLng(nRGB.Green), CLng(nRGB.Blue))
            'mOffscreen.SetPixel x, (y - 8), mlGrayScale(miPixelData(x + ((y - 8) * dWidth)))
            
        Next X
        If Not Progress Is Nothing Then
            Progress.value = Progress.value + 1
        End If
    Next Y

    'Dim dummy As Long
    'If dWidth = 256 And dHeight = 256 Then
    '    dummy = 65535
    'End If
    
    'ReDim Preserve miPixelData(dWidth * dHeight)
    
    mOffscreen.width = dWidth
    mOffscreen.height = dHeight
    
    BitBlt BMPOffscreen.hdc, 0, 0, BMPOffscreen.width, BMPOffscreen.height - 8, BMPOffscreen.hdc, 0, 8, vbSrcCopy
    
    BMPOffscreen.width = dWidth
    BMPOffscreen.height = dHeight
     
    mOffscreen.Create CLng(dWidth), CLng(dHeight)
    
    'match pals
    Dim iPalMap() As Byte
    mMatchPal BMPOffscreen, PalMap, Progress, Status
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:GetBitFromBMP Error"
End Sub


Public Property Let GrayScale(Index As Integer, lNewValue As Long)

    mlGrayScale(Index) = lNewValue

End Property

Public Property Get GrayScale(Index As Integer) As Long

    GrayScale = mlGrayScale(Index)

End Property



Private Sub mMatchPal(BMPOffscreen As clsOffscreen, iPalMap() As Byte, Optional Progress As ProgressBar, Optional Status As StatusBar)

    On Error GoTo HandleErrors

    Dim bitColor As New clsRGB
    Dim i As Integer
    Dim xTile As Integer
    Dim yTile As Integer
    Dim xPixel As Byte
    Dim yPixel As Byte
    Dim pal As Byte
    Dim colr As Byte
    Dim highPalMatch(7) As Byte
    Dim highPal As Byte
    Dim palMatch As Byte
    
    Screen.MousePointer = vbHourglass
    
    ReDim iPalMap((Offscreen.width \ 8), (Offscreen.height \ 8))
    
    If Not Status Is Nothing Then
        Status.Panels(1).Text = "Matching palettes..."
    End If
    
    If Not Progress Is Nothing Then
        Progress.value = 0
        Progress.Max = 40 + (BMPOffscreen.height \ 8)
    End If
    
    Dim oPalsFromPal() As New clsPalFromPal
    Dim oPalsFromTile() As clsPalFromTile
    Dim oPools() As clsPool
    
    GetPoolsFromPal mGBPalette, oPools, oPalsFromPal, Progress
    GetPalsFromTiles BMPOffscreen, oPools, oPalsFromTile, Progress
    GetPalMapFromPools oPools, oPalsFromPal, oPalsFromTile, iPalMap, Progress
    
    mRenderBitmap BMPOffscreen, iPalMap, Progress, Status

''debug print
'Dim x As Integer
'Dim y As Integer
'Dim nFilenum As Integer
'nFilenum = FreeFile
'Open "c:\windows\desktop\palmap.txt" For Output As #nFilenum
'For y = 0 To UBound(iPalMap, 2) - 1
'    For x = 0 To UBound(iPalMap, 1) - 1
'        Print #nFilenum, iPalMap(x, y);
'    Next x
'    Print #nFilenum, ""
'Next y
'Close #nFilenum
 
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:mMatchPal Error"
End Sub


Private Sub mRenderBitmap(BMPOffscreen As clsOffscreen, iPalMap() As Byte, Optional Progress As ProgressBar, Optional Status As StatusBar)

    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim PixelX As Byte
    Dim PixelY As Byte
    Dim pixelRGB As New clsRGB
    Dim PalID As Byte
    Dim bestColor As Byte

    On Error GoTo HandleErrors

    If Not Progress Is Nothing Then
        Progress.value = 0
        Progress.Max = (Offscreen.height \ 8) * (Offscreen.width \ 8)
    End If
    
    If Not Status Is Nothing Then
        Status.Panels(1).Text = "Rendering bitmap..."
    End If

    For Y = 0 To (Offscreen.height \ 8) - 1
        For X = 0 To (Offscreen.width \ 8) - 1
            For PixelY = 0 To 7
                For PixelX = 0 To 7

                    Set pixelRGB = GetRGBFromLong(GetPixel(BMPOffscreen.hdc, (X * 8) + PixelX, (Y * 8) + PixelY))
                    With pixelRGB
                        .Red = (.Red \ 8) And &H1F
                        .Green = (.Green \ 8) And &H1F
                        .Blue = (.Blue \ 8) And &H1F
                    End With
                    
                    If iPalMap(X, Y) = 0 Then
                        AddTileGetterError
                        TileGetterErrors(TileGetterErrorCount).nType = NoPaletteMatch
                        TileGetterErrors(TileGetterErrorCount).X = X
                        TileGetterErrors(TileGetterErrorCount).Y = Y
                        GoTo NextTile
                    End If
                    PalID = iPalMap(X, Y) - 1

                    bestColor = 0
                    For i = 1 To 4
                        If GBPalette.Colors(i + (PalID * 4)).Red = pixelRGB.Red And GBPalette.Colors(i + (PalID * 4)).Green = pixelRGB.Green And GBPalette.Colors(i + (PalID * 4)).Blue = pixelRGB.Blue Then
                            bestColor = i - 1
                            Exit For
                        End If
                    Next i

                    PixelData((X * 8) + PixelX, (Y * 8) + PixelY) = bestColor

                Next PixelX
            Next PixelY
NextTile:
            If Not Progress Is Nothing Then
                Progress.value = Progress.value + 1
            End If
        Next X
    Next Y

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:mRenderBitmap Error"
End Sub


Public Function GetGrayFromRGB(Red As Long, Green As Long, Blue As Long) As Byte
    
'***************************************************************************
'   Map an RGB color to a 4-intensity grayscale
'***************************************************************************
    
    Dim intensity As Long
    
    intensity = (Red + Green + Blue) / 3
    
    If intensity < 64 Then
        GetGrayFromRGB = 0
        Exit Function
    End If

    If intensity < 128 Then
        GetGrayFromRGB = 1
        Exit Function
    End If
    
    If intensity < 192 Then
        GetGrayFromRGB = 2
        Exit Function
    End If
    
    GetGrayFromRGB = 3

End Function

Public Sub PackToBin(ByVal sFilename As String)

 '   s16         x;
 '   u8          xShift;
 '   u16         y;
 '   u16         xTile;
 '   u16         yTile;
 '   u8          lo;
 '   u8          hi;
 '   u8          color;

'   if ( !BinFile_Create ( &binFile, fileName )) return 0;'

'    for ( yTile = 0; yTile < (bitmap.Height>>3); yTile++ )
'        for ( xTile = 0; xTile < (bitmap.Width>>3); xTile++ )
'            for ( y = 0; y < 8; y++ )
'                {
'                lo = 0;
'                hi = 0;
'                for ( x = 0;  x < 8; x++ )
'                    {
'                    color = (u8)(pixels[(xTile<<3)+x][(yTile<<3)+y] & 0x03);
'                    xShift = (7-x);
'                    lo += ((color&0x01)<<xShift);
'                    hi += (((color&0x02)>>1)<<xShift);
'                    }
'                BinFile_Write ( &binFile, &lo, sizeof ( u8 ));
'                BinFile_Write ( &binFile, &hi, sizeof ( u8 ));
'                }

'    BinFile_Close ( &binFile );

'    return 1;
'    }
    
    If sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Binary Access Write As #nFilenum
    
    Dim X As Integer
    Dim Y As Integer
    Dim xShift As Byte
    Dim yShift As Byte
    Dim xTile As Integer
    Dim yTile As Integer
    Dim lo As Byte
    Dim hi As Byte
    Dim color As Byte
    Dim nCount As Integer

    For yTile = 0 To (height \ 8) - 1
        For xTile = 0 To (width \ 8) - 1
            For Y = 0 To 7
                lo = 0
                hi = 0
                For X = 0 To 7
                    color = miPixelData(((xTile * 8) + X) + (((yTile * 8) + Y) * width))
                    xShift = 7 - X
                    lo = lo + BitShiftLeft((color And &H1), xShift)
                    hi = hi + BitShiftLeft(BitShiftRight((color And &H2), 1), xShift)
                Next X
                Put #nFilenum, , lo
                Put #nFilenum, , hi
            Next Y
            
            If mnTileCount > 0 Then
                nCount = nCount + 1
                If nCount = mnTileCount Then
                    GoTo rEnd
                End If
            End If
            
        Next xTile
    Next yTile

rEnd:
    Close #nFilenum
End Sub

Public Sub RenderPixels()

    On Error GoTo HandleErrors

    Dim xTile As Long
    Dim yTile As Long
    Dim xPixel As Integer
    Dim yPixel As Integer
    Dim nCount As Integer
    
    For yTile = 0 To 31 '(height \ 8) - 1
        For xTile = 0 To (width \ 8) - 1
        
            If Not mbClipping Then
                If yTile > (height \ 8) - 1 Then
                    Exit Sub
                End If
            End If
        
            For yPixel = 0 To 7
                For xPixel = 0 To 7
        
                    mOffscreen.SetPixel (xTile * 8) + xPixel, (yTile * 8) + yPixel, mlGrayScale(PixelData((xTile * 8) + xPixel, (yTile * 8) + yPixel))
                    
                Next xPixel
            Next yPixel
        
            If mnTileCount > 0 Then
                nCount = nCount + 1
                If (nCount = mnTileCount) Then
                    mOffscreen.height = (yTile + 1) * 8
                    Exit Sub
                End If
            End If
        
        Next xTile
    Next yTile

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:RenderPixels Error"
End Sub

Public Sub RenderPixels1()

    On Error GoTo HandleErrors

    Dim X As Long
    Dim Y As Long
    Dim oRGB As New clsRGB
    
    For Y = 0 To height - 1
        For X = 0 To width - 1
        
            Set oRGB = GetRGBFromLong(mOffscreen.GetPixel(X, Y))
        
            PixelData(X, Y) = GetGrayFromRGB(oRGB.Red, oRGB.Green, oRGB.Blue)
        
        Next X
    Next Y

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:RenderPixels1 Error"
End Sub

Public Sub ResizePixelData(nWidth As Integer, nHeight As Integer)

    ReDim miPixelData(nWidth * nHeight)
    
End Sub

Public Property Get ResourceCacheObject() As clsResourceCacheObject

    Set ResourceCacheObject = mResourceCacheObject

End Property

Public Property Let TileCount(nNewValue As Integer)

    mnTileCount = nNewValue

End Property

Public Property Get TileCount() As Integer

    TileCount = mnTileCount

End Property


Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient

End Sub

Private Sub Class_Initialize()
    
'***************************************************************************
'   Initialize 4-intensity grayscale display values
'***************************************************************************
    
    mlGrayScale(0) = RGB(0, 7, 0)
    mlGrayScale(1) = RGB(47, 79, 47)
    mlGrayScale(2) = RGB(131, 167, 131)
    mlGrayScale(3) = RGB(255, 255, 255)
    
End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Property Set intResource_CacheObject(oNewValue As clsResourceCacheObject)
    
    Set mResourceCacheObject = oNewValue
    
End Property

Private Sub intResource_Delete()
    
'***************************************************************************
'   Release memory when the class is terminated
'***************************************************************************

    If Not mOffscreen Is Nothing Then
        mOffscreen.Delete
    End If
    
End Sub

Private Function intResource_Unpack(nFilenum As Integer) As Boolean
       
'***************************************************************************
'   Unpack GB bitmap data from a currently open file represented by nFilenum
'***************************************************************************
       
    intResource_Unpack = True
       
'Setup error handling
    On Error GoTo HandleErrors
    
    Dim X As Long
    Dim Y As Long
    Dim lWidth As Long
    Dim lHeight As Long
    
'Get width and height of the bitmap
    Get #nFilenum, , lWidth
    Get #nFilenum, , lHeight
    
    mOffscreen.Create lWidth, lHeight
        
    ReDim miPixelData(width * height)
    
'Unpack pixel data and store it in the GBBitmap class variable
    For Y = 0 To height - 1
        For X = 0 To width - 1
            Get #nFilenum, , miPixelData(X + (Y * width))
            mOffscreen.SetPixel X, Y, mlGrayScale(miPixelData(X + (Y * width)))
        Next X
    Next Y

    Get #nFilenum, , mnTileCount
    
    If Not gbBatchExport Then
        RenderPixels
    End If

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:Unpack Error"
    intResource_Unpack = False
End Function

Public Sub UnpackBitmapFromBMP(ByVal sFilename As String)

'***************************************************************************
'   Unpack a Windows Bitmap and save it in the GBBitmap class variable
'***************************************************************************

    mOffscreen.CreateBitmapFromBMP sFilename
    
    Dim X As Long
    Dim Y As Long
    Dim nRGB As New clsRGB
    
    ReDim miPixelData(width * height)
    
    For Y = 0 To height - 1
        For X = 0 To width - 1
            'DoEvents
            Set nRGB = GetRGBFromLong(mOffscreen.GetPixel(X, Y))
            miPixelData(X + (Y * width)) = GetGrayFromRGB(CLng(nRGB.Red), CLng(nRGB.Green), CLng(nRGB.Blue))
            mOffscreen.SetPixel X, Y, mlGrayScale(miPixelData(X + (Y * width)))
        Next X
    Next Y

End Sub

Public Sub UnpackBitmapFromAdvBMP(ByVal sFilename As String, Optional Progress As ProgressBar, Optional Status As StatusBar, Optional StatusText As String)

'***************************************************************************
'   Unpack a Windows Bitmap and save it in the GBBitmap class variable
'***************************************************************************

    On Error GoTo HandleErrors

    mOffscreen.CreateBitmapFromBMP sFilename
    
    Dim X As Long
    Dim Y As Long
    Dim flag As Boolean
    Dim nRGB As New clsRGB
    Dim lCropColor As Long
    
    'get palette data from bitmap
    For X = 0 To 31
        Set nRGB = GetRGBFromLong(mOffscreen.GetPixel(X, 0))
        mGBPalette.Colors(X + 1).Red = (nRGB.Red \ 8) And &H1F
        mGBPalette.Colors(X + 1).Green = (nRGB.Green \ 8) And &H1F
        mGBPalette.Colors(X + 1).Blue = (nRGB.Blue \ 8) And &H1F
    Next X
    
    'get crop color
    lCropColor = mOffscreen.GetPixel(32, 0)
    
    For X = 0 To width - 1
        mOffscreen.SetPixel X, height - 1, lCropColor
    Next X
    
    For Y = 0 To height - 1
        mOffscreen.SetPixel width - 1, Y, lCropColor
    Next Y
    
    'get bitmap data
    ReDim miPixelData(width * height)
    
    Dim dWidth As Integer
    Dim dHeight As Integer
    
    dWidth = width
    dHeight = height
    
    If Not Progress Is Nothing Then
        Progress.value = 0
        Progress.Max = height
    End If
    
    If Not Status Is Nothing Then
        If StatusText = "" Then
            Status.Panels(1).Text = "Unpacking bitmap..."
        Else
            Status.Panels(1).Text = StatusText
        End If
    End If
    
    For Y = 8 To height - 1
        If mOffscreen.GetPixel(0, Y) = lCropColor Then
            dHeight = Y - 8
            If Not Progress Is Nothing Then
                Do Until Progress.value = Progress.Max
                    Progress.value = Progress.value + 1
                Loop
            End If
            Exit For
        End If
        For X = 0 To width - 1
            'DoEvents
            If (mOffscreen.GetPixel(X, 8) = lCropColor) And flag = False Then
                flag = True
                dWidth = X
                Exit For
            End If
            Set nRGB = GetRGBFromLong(mOffscreen.GetPixel(X, Y))
            miPixelData(X + ((Y) * width)) = GetGrayFromRGB(CLng(nRGB.Red), CLng(nRGB.Green), CLng(nRGB.Blue))
            mOffscreen.SetPixel X, (Y - 8), mlGrayScale(miPixelData(X + ((Y) * width)))
        
        Next X
        If Not Progress Is Nothing Then
            Progress.value = Progress.value + 1
        End If
    Next Y

    Dim dummy As Long
    If dWidth = 256 And dHeight = 256 Then
        dummy = 65535
    End If
    ReDim Preserve miPixelData(dummy)
    
    mOffscreen.width = dWidth
    mOffscreen.height = dHeight

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:UnpackBitmapFromAdvBMP Error"
End Sub

Private Sub intResource_Pack(nFilenum As Integer)

    On Error GoTo HandleErrors

'***************************************************************************
'   Pack the bitmap into a .bit file
'***************************************************************************

    Dim X As Integer
    Dim Y As Integer
    
'Write file type
    Put #nFilenum, , CByte(GB_BITMAP)
    
'Write bitmap dimensions
    Put #nFilenum, , width
    Put #nFilenum, , height
    
'Write pixel data
    For Y = 0 To height - 1
        For X = 0 To width - 1
            Put #nFilenum, , miPixelData(X + (Y * width))
        Next X
    Next Y
    
    Put #nFilenum, , mnTileCount

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBitmap:Pack Error"
End Sub

Private Property Get intResource_ResourceType() As GB_FILETYPES
    
'***************************************************************************
'   Return the type of resource this is
'***************************************************************************
    
    intResource_ResourceType = GB_BITMAP
    
End Property

Private Sub intResource_UpdateClients(iClient As intResourceClient)
    
'***************************************************************************
'   Update all clients of this resource
'***************************************************************************
    
    If Not mResourceCacheObject Is Nothing Then
        mResourceCacheObject.UpdateClients iClient
    End If
    
End Sub

Public Property Get Offscreen() As clsOffscreen
    
    Set Offscreen = mOffscreen
    
End Property

Public Property Set Offscreen(oNewValue As clsOffscreen)
    
    Set mOffscreen = oNewValue
    
End Property

Public Property Get PixelData(X As Long, Y As Long) As Byte
    
    PixelData = miPixelData(X + (Y * width))
    
End Property

Public Property Let PixelData(X As Long, Y As Long, iNewValue As Byte)
    
    miPixelData(X + (Y * width)) = iNewValue
    
End Property

Public Sub BlitWithPal(hDestDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, ByVal pal As clsGBPalette, ByVal PalID As Integer, Optional XFlip As Integer, Optional YFlip As Integer, Optional BlitTransparent As Boolean)
    
'***************************************************************************
'   Do a blit of the bitmap using a specified palette and palette index
'***************************************************************************
    
    Dim nx As Integer
    Dim ny As Integer
    Dim startX As Integer
    Dim startY As Integer
    Dim endX As Integer
    Dim endY As Integer
    Dim incX As Integer
    Dim incY As Integer
    Dim countX As Integer
    Dim countY As Integer
    Dim Index As Integer
    Dim pdata As Byte
    
    If XFlip = 0 Then
        startX = 0
        endX = nWidth - 1
        incX = 1
    Else
        startX = nWidth - 1
        endX = 0
        incX = -1
    End If
    
    If YFlip = 0 Then
        startY = 0
        endY = nHeight - 1
        incY = 1
    Else
        startY = nHeight - 1
        endY = 0
        incY = -1
    End If
    
    For ny = startY To endY Step incY
        For nx = startX To endX Step incX
            If ((nx + xSrc) < width) And ((ny + ySrc) < height) Then
                pdata = miPixelData((nx + xSrc) + ((ny + ySrc) * width))
                If PalID < 8 Then
                    Index = (PalID * 4) + pdata + 1
                Else
                    Index = pdata + 1
                End If
                If BlitTransparent Then
                    If pdata <> 0 Then
                        SetPixel hDestDC, X + countX, Y + countY, RGB(pal.Colors(Index).Red * 8, pal.Colors(Index).Green * 8, pal.Colors(Index).Blue * 8)
                    End If
                Else
                    SetPixel hDestDC, X + countX, Y + countY, RGB(pal.Colors(Index).Red * 8, pal.Colors(Index).Green * 8, pal.Colors(Index).Blue * 8)
                End If
            End If
            countX = countX + 1
        Next nx
        countX = 0
        countY = countY + 1
    Next ny

End Sub

Public Sub BlitWithPalToMask(bitmapDC As Long, maskDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, ByVal pal As clsGBPalette, ByVal PalID As Integer)
    
'***************************************************************************
'   Do a blit of the bitmap using a specified palette and palette index and create a mask
'***************************************************************************
    
    Dim nx As Integer
    Dim ny As Integer
    Dim Index As Integer
    Dim pdata As Byte
    
    For ny = 0 To nHeight - 1
        For nx = 0 To nWidth - 1
            If ((nx + xSrc) < width) And ((ny + ySrc) < height) Then
                pdata = miPixelData((nx + xSrc) + ((ny + ySrc) * width))
                If pdata = 0 Then
                    SetPixel bitmapDC, X + nx, Y + ny, vbBlack
                    SetPixel maskDC, X + nx, Y + ny, vbWhite
                Else
                    Index = (PalID * 4) + pdata + 1
                    SetPixel bitmapDC, X + nx, Y + ny, RGB(pal.Colors(Index).Red * 8, pal.Colors(Index).Green * 8, pal.Colors(Index).Blue * 8)
                    SetPixel maskDC, X + nx, Y + ny, vbBlack
                End If
            End If
        Next nx
    Next ny

End Sub

Public Property Get width() As Long

    width = mOffscreen.width

End Property

Public Property Let width(lNewValue As Long)
    
    mOffscreen.width = lNewValue
    
End Property

Public Property Get height() As Long

    height = mOffscreen.height

End Property

Public Property Let height(lNewValue As Long)
    
    mOffscreen.height = lNewValue
    
End Property

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

'***************************************************************************
'   Update clients of the resource
'***************************************************************************
    
    intResource_UpdateClients Nothing

End Sub


