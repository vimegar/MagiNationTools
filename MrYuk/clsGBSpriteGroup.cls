VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBSpriteGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements intResource
Implements intResourceClient

Private mSprites() As New clsSpriteTileGroup
Private mOffscreen As New clsOffscreen

Private msVRAMFile As String
Private msPaletteFile As String
Private mGBVRAM As New clsGBVRAM
Private mGBPalette As New clsGBPalette
Private msParentPath As String

Private mResourceCacheObject As clsResourceCacheObject

Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property

Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Public Sub Delete()

    intResource_Delete

End Sub


Public Property Set GBPalette(oNewValue As clsGBPalette)

    Set mGBPalette = oNewValue

End Property

Public Property Get GBPalette() As clsGBPalette

    Set GBPalette = mGBPalette

End Property

Public Property Set GBVRAM(oNewValue As clsGBVRAM)

    Set mGBVRAM = oNewValue

End Property

Public Property Get GBVRAM() As clsGBVRAM

    Set GBVRAM = mGBVRAM

End Property

Public Property Set Offscreen(oNewValue As clsOffscreen)
    
    Set mOffscreen = oNewValue

End Property

Public Property Get Offscreen() As clsOffscreen

    Set Offscreen = mOffscreen

End Property

Public Sub PackToBin(ByVal sFilename As String)

    If sFilename = "" Then
        Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    Dim strX As String
    Dim strY As String
    Dim strTab As String
    Dim strTile As String
    Dim strFlags As String
    Dim spung As Byte
    Dim nFilenum As Integer
    nFilenum = FreeFile

    Open sFilename For Output As #nFilenum
    
    strTab = Chr(vbKeyTab)
    
    Dim nFilesize As Integer
    nFilesize = 0
    For i = 1 To SpriteCount
        For j = mSprites(i).TileCount To 1 Step -1
            nFilesize = nFilesize + 4
        Next j
    Next i
    
    nFilesize = nFilesize + 1
    
    Print #nFilenum, ";Size:" & CStr(nFilesize)
    Print #nFilenum, ";********************************"
    Print #nFilenum, "; " & UCase$(GetTruncFilename(sFilename))
    Print #nFilenum, ";********************************"
    Print #nFilenum, ";" & strTab & "Author:" & strTab & "Mr. Yuk"
    Print #nFilenum, ";" & strTab & "(c)2000" & strTab & "Interactive Imagination"
    Print #nFilenum, ";" & strTab & "All rights reserved"
    Print #nFilenum, ""
    Print #nFilenum, ";********************************"
    
    For i = 1 To SpriteCount
        
        Print #nFilenum, "?" & mSprites(i).Name
        
        'For j = 1 To mSprites(i).TileCount
        For j = mSprites(i).TileCount To 1 Step -1
        
            With mSprites(i).Tiles(j)
                
                If .XOffset - 128 < 0 Then
                    strX = "$" & Hex(((Not (Abs(.XOffset - 128) And &HFF)) + 1) And &HFF) & "," & strTab
                Else
                    strX = "$" & Hex(.XOffset - 128) & "," & strTab
                End If
                
                If .yOffset - 128 < 0 Then
                    strY = "$" & Hex(((Not (Abs(.yOffset - 128) And &HFF)) + 1) And &HFF) & "," & strTab
                Else
                    strY = "$" & Hex(.yOffset - 128) & "," & strTab
                End If
                
                strTile = "$" & Hex(.TileID And &HFF) & "," & strTab
                
                If Len(strTile) < 5 Then
                    strTile = "$0" & Mid$(strTile, 2)
                End If
                
                If Len(strX) < 5 Then
                    strX = "$0" & Mid$(strX, 2)
                End If
                
                If Len(strY) < 5 Then
                    strY = "$0" & Mid$(strY, 2)
                End If
                
                spung = CByte(.PalID)
                spung = CByte(spung + .Bank * 8)
                spung = CByte(spung + .XFlip * 32)
                spung = CByte(spung + .YFlip * 64)
                spung = CByte(spung + .Priority * 128)

                strFlags = "$" & Hex(spung And &HFF)

                If Len(strFlags) < 3 Then
                    strFlags = "$0" & Mid$(strFlags, 2)
                End If

                Print #nFilenum, strTab & "DB" & strTab;
                Print #nFilenum, strY & strX & strTile & strFlags
            
            End With
        
        Next j
        
        Print #nFilenum, strTab & "DB" & strTab & "$80"
        Print #nFilenum, ""
        Print #nFilenum, ";********************************"
    
    Next i
    
    Print #nFilenum, strTab & "END"
    Print #nFilenum, ";********************************"
    
    Close #nFilenum
    
End Sub

Public Property Let sPaletteFile(sNewValue As String)

    msPaletteFile = sNewValue

End Property

Public Property Get sPaletteFile() As String

    sPaletteFile = msPaletteFile

End Property

Public Property Set Sprites(Index As Integer, oNewValue As clsSpriteTileGroup)

    Set mSprites(Index) = oNewValue

End Property

Public Property Get Sprites(Index As Integer) As clsSpriteTileGroup

    Set Sprites = mSprites(Index)

End Property

Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient

End Sub

Private Sub Class_Initialize()

    ReDim mSprites(0)

End Sub

Public Property Get SpriteCount() As Integer

    SpriteCount = UBound(mSprites)

End Property
Public Sub AddSprite()

'***************************************************************************
'   Add a sprite
'***************************************************************************

    ReDim Preserve mSprites(UBound(mSprites) + 1)

End Sub

Public Sub DeleteSprite(Index As Integer)

'***************************************************************************
'   Delete a sprite group based on its index
'***************************************************************************

    Dim i As Integer
    
    For i = Index To UBound(mSprites) - 1
        Set mSprites(i) = mSprites(i + 1)
    Next i
    
    ReDim Preserve mSprites(UBound(mSprites) - 1)

End Sub

Public Property Let sVRAMFile(sNewValue As String)

    msVRAMFile = sNewValue

End Property

Public Property Get sVRAMFile() As String

    sVRAMFile = msVRAMFile
    
End Property

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Property Set intResource_CacheObject(RHS As clsResourceCacheObject)

    Set mResourceCacheObject = RHS

End Property

Private Sub intResource_Delete()

    If Not mOffscreen Is Nothing Then
        mOffscreen.Delete
    End If
    
End Sub

Private Sub intResource_Pack(nFilenum As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim s As Integer
    
    Put #nFilenum, , CByte(GB_SPRITEGROUP)
    
    s = Len(msVRAMFile)
    Put #nFilenum, , s
    Put #nFilenum, , msVRAMFile
    
    s = Len(msPaletteFile)
    Put #nFilenum, , s
    Put #nFilenum, , msPaletteFile
    
    Put #nFilenum, , CInt(SpriteCount)

    For i = 1 To SpriteCount
    
        s = Len(mSprites(i).Name)
        Put #nFilenum, , s
        Put #nFilenum, , mSprites(i).Name
        
        Put #nFilenum, , CInt(mSprites(i).TileCount)
        
        For j = 1 To mSprites(i).TileCount
            Put #nFilenum, , CByte(mSprites(i).Tiles(j).TileID)
            Put #nFilenum, , CByte(mSprites(i).Tiles(j).Bank)
            Put #nFilenum, , CByte(mSprites(i).Tiles(j).PalID)
            Put #nFilenum, , CByte(mSprites(i).Tiles(j).XFlip)
            Put #nFilenum, , CByte(mSprites(i).Tiles(j).YFlip)
            Put #nFilenum, , CInt(mSprites(i).Tiles(j).XOffset)
            Put #nFilenum, , CInt(mSprites(i).Tiles(j).yOffset)
            Put #nFilenum, , CInt(mSprites(i).Tiles(j).BitmapFragmentIndex)
        Next j
        
    Next i

End Sub

Private Property Get intResource_ResourceType() As GB_FILETYPES
    
    intResource_ResourceType = GB_SPRITEGROUP
    
End Property

Private Function intResource_Unpack(nFilenum As Integer) As Boolean

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim j As Integer
    Dim s As Integer
    Dim ds As String
    Dim n As Integer
    Dim d As Integer
    Dim b As Byte
    Dim count As Integer
    
    intResource_Unpack = True
    
    Get #nFilenum, , s
    msVRAMFile = String(s, " ")
    Get #nFilenum, , msVRAMFile
    
    msVRAMFile = msParentPath & "\VRAMs\" & GetTruncFilename(msVRAMFile)
    
    If Not gbBatchExport Then
        
        gbOpeningChild = True
        Set mGBVRAM = gResourceCache.GetResourceFromFile(msVRAMFile, Me)
        gbOpeningChild = False
        
        If mGBVRAM.CacheObject Is Nothing Then
            intResource_Unpack = False
            Exit Function
        End If
    
    End If
    
    Get #nFilenum, , s
    msPaletteFile = String(s, " ")
    Get #nFilenum, , msPaletteFile
    
    msPaletteFile = msParentPath & "\Palettes\" & GetTruncFilename(msPaletteFile)
    
    If Not gbBatchExport Then
        
        gbOpeningChild = True
        Set mGBPalette = gResourceCache.GetResourceFromFile(msPaletteFile, Me)
        gbOpeningChild = False
        
        If mGBPalette.CacheObject Is Nothing Then
            intResource_Unpack = False
            Exit Function
        End If
    End If
    
    Get #nFilenum, , n
    For count = 1 To n
        AddSprite
    Next count

    For i = 1 To SpriteCount
    
        Get #nFilenum, , s
        ds = String(s, " ")
        Get #nFilenum, , ds
        mSprites(i).Name = ds
    
        Get #nFilenum, , n
        For count = 1 To n
            mSprites(i).AddTile
        Next count
        
        For j = 1 To mSprites(i).TileCount
            DoEvents
            
            Get #nFilenum, , b
            mSprites(i).Tiles(j).TileID = b
        
            Get #nFilenum, , b
            mSprites(i).Tiles(j).Bank = b
        
            Get #nFilenum, , b
            mSprites(i).Tiles(j).PalID = b
        
            Get #nFilenum, , b
            mSprites(i).Tiles(j).XFlip = b
        
            Get #nFilenum, , b
            mSprites(i).Tiles(j).YFlip = b
        
            Get #nFilenum, , d
            mSprites(i).Tiles(j).XOffset = d
        
            Get #nFilenum, , d
            mSprites(i).Tiles(j).yOffset = d
            
            Get #nFilenum, , d
            mSprites(i).Tiles(j).BitmapFragmentIndex = d
        Next j
        
    Next i

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBSpriteGroup:intResource_Unpack Error"
End Function

Private Sub intResource_UpdateClients(iClient As intResourceClient)

    If Not mResourceCacheObject Is Nothing Then
        mResourceCacheObject.UpdateClients iClient
    End If

End Sub

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    intResource_UpdateClients Nothing

End Sub

