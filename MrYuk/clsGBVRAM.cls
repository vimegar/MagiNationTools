VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBVRAM"
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

    Private Type tVRAMEntry
        Filename As String
        BaseAddress As Long
        Bank As Byte
        Bitmap As clsGBBitmap
    End Type

    Private mVRAMEntries() As tVRAMEntry
    Private mOffscreenBank0 As New clsOffscreen
    Private mOffscreenBank1 As New clsOffscreen
    Private mBitmapFragments(384, 1) As New clsBitmapFragment
    Private msParentPath As String
    
'***************************************************************************
'   Resource cache object pointers
'***************************************************************************

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

Public Function GetBitFragIDFromVRAMAddr(lAddr As Long) As Integer

    If lAddr < 32768 Then
        Exit Function
    End If
    
    Dim point As New clsPoint

    GetPointFromVRAMAddress point, lAddr
    
    point.X = point.X \ 8
    point.Y = point.Y \ 8
    
    GetBitFragIDFromVRAMAddr = point.X + (point.Y * 16) + 1
    
End Function


Public Sub GetPointFromVRAMAddress(gc As clsPoint, addr As Long)
    
'***************************************************************************
'   Returns a point class based on a VRAM address
'***************************************************************************
    
    If addr < 32768 Then
        gc.X = 0
        gc.Y = 0
    Else
        gc.X = (addr And &HF0) / 2
        gc.Y = ((addr - 32768) And &HFF00) / 32
    End If
    
End Sub

Public Function GetVRAMAddressFromPoint(gc As clsPoint) As Long
    
    GetVRAMAddressFromPoint = 32768 + (gc.X * 16) + (gc.Y * 256)
    
End Function

Public Sub EnumBitmapFragments()

'***************************************************************************
'   Update the bitmap fragments using the VRAM entry information
'***************************************************************************
    
    On Error Resume Next
    
    Dim Entry As Integer
    Dim gcSrc As New clsPoint
    Dim gcDest As New clsPoint
    Dim copyOK As Boolean
    Dim i As Integer
    Dim j As Integer
    
'Get the most current resource
    For i = 1 To gResourceCache.TotalCacheObjects
        For j = 1 To UBound(mVRAMEntries)
            If gResourceCache.CacheObjectList(i).sFilename = mVRAMEntries(j).Filename Then
                Set mVRAMEntries(j).Bitmap = gResourceCache.CacheObjectList(i).Data
            End If
        Next j
    Next i
    
'Deallocate bitmap fragments
    For i = 1 To 384
        For j = 0 To 1
            mBitmapFragments(i, j).X = 0
            mBitmapFragments(i, j).Y = 0
            Set mBitmapFragments(i, j).GBBitmap = Nothing
        Next j
    Next i
    
    Dim nCount As Integer
    
    For Entry = 1 To VRAMEntryCount
        
        copyOK = True
        GetPointFromVRAMAddress gcDest, VRAMEntryBaseAddress(Entry)
        GetPointFromVRAMAddress gcSrc, 0
        
        Dim startX As Integer
        Dim startY As Integer
        startX = gcDest.X
        startY = gcDest.Y
        
        nCount = 0
        
        While copyOK
            
            If VRAMEntryBitmap(Entry).TileCount > 0 Then
                If nCount = VRAMEntryBitmap(Entry).TileCount Then
                    GoTo NextEntry
                End If
                nCount = nCount + 1
            End If
            
        'Set scratch variable equal to the parameters of the current VRAM entry
            With mBitmapFragments(gcDest.X \ 8 + ((gcDest.Y \ 8) * 16) + 1, VRAMEntryBank(Entry))
                .X = gcSrc.X
                .Y = gcSrc.Y
                Set .GBBitmap = VRAMEntryBitmap(Entry)
            End With
            
        'Handle exiting logic
            
            Dim tempSrcY As Integer
            tempSrcY = gcSrc.Y
            If Not GridCursorNext(gcSrc, VRAMEntryBitmap(Entry).Offscreen.width, VRAMEntryBitmap(Entry).Offscreen.height) Then
                copyOK = False
            End If
            
            If gcDest.X + 8 >= 128 Or gcDest.X + 8 >= VRAMEntryBitmap(Entry).Offscreen.width + startX Then
                If tempSrcY >= gcSrc.Y Then
                    gcSrc.X = 0
                    gcSrc.Y = gcSrc.Y + 8
                    nCount = nCount + (((startX + VRAMEntryBitmap(Entry).Offscreen.width) - 128) \ 8)
                End If
                gcDest.X = startX
                gcDest.Y = gcDest.Y + 8
                If gcDest.Y >= 192 Then
                    copyOK = False
                End If
            Else
                If Not GridCursorNext(gcDest, VRAMEntryBitmap(Entry).Offscreen.width + startX, VRAMEntryBitmap(Entry).Offscreen.height + startY) Then
                    copyOK = False
                End If
            End If
        
        Wend
    
NextEntry:
    Next Entry

End Sub

Public Function GetTileCount() As Integer

    Dim nCount As Integer
    Dim addr As Long
    Dim j As Integer
    Dim nBank As Integer
    Dim ipnt As New clsPoint
    Dim jpnt As New clsPoint
    Dim d As RECT
    
    For nBank = 0 To 1
        For addr = 32768 To 40959 Step 16
            
            GetPointFromVRAMAddress ipnt, addr
            
            For j = 1 To VRAMEntryCount
                
                If VRAMEntryBank(j) = nBank Then
                    
                    GetPointFromVRAMAddress jpnt, VRAMEntryBaseAddress(j)
                    d.Left = jpnt.X
                    d.Top = jpnt.Y
                    d.Right = jpnt.X + VRAMEntryBitmap(j).width - 1
                    d.Bottom = jpnt.Y + VRAMEntryBitmap(j).height - 1
                    
                    If ipnt.X >= d.Left And ipnt.X <= d.Right And ipnt.Y >= d.Top And ipnt.Y <= d.Bottom Then
                        nCount = nCount + 1
                    End If
                
                End If
                
            Next j
            
        Next addr
    Next nBank
    
    GetTileCount = nCount

End Function

Public Function GridCursorNext(gc As clsPoint, width As Integer, height As Integer) As Boolean

'***************************************************************************
'   GridCursor manipulation routine used for display purposes
'***************************************************************************

    GridCursorNext = True

    gc.X = gc.X + 8
    
    If gc.X >= width Then
        gc.X = 0
        gc.Y = gc.Y + 8
        
        If gc.Y >= height Then
            GridCursorNext = False
        End If
    End If
    
End Function

Public Property Get OffscreenBank0() As clsOffscreen

    Set OffscreenBank0 = mOffscreenBank0

End Property

Public Property Set OffscreenBank0(oNewValue As clsOffscreen)

    Set mOffscreenBank0 = oNewValue

End Property

Public Property Get OffscreenBank1() As clsOffscreen

    Set OffscreenBank1 = mOffscreenBank1

End Property

Public Property Set OffscreenBank1(oNewValue As clsOffscreen)

    Set mOffscreenBank1 = oNewValue

End Property

Public Property Set BitmapFragments(Index As Integer, Bank As Integer, oNewValue As clsBitmapFragment)

    Set mBitmapFragments(Index, Bank) = oNewValue

End Property

Public Property Get BitmapFragments(Index As Integer, Bank As Integer) As clsBitmapFragment

    Set BitmapFragments = mBitmapFragments(Index, Bank)

End Property

Public Sub Delete()

    intResource_Delete

End Sub

Public Sub DeleteVRAMEntry(Index As Integer)

'***************************************************************************
'   Delete a VRAM entry based on its index
'***************************************************************************

    Dim i As Integer
    
    For i = Index To UBound(mVRAMEntries) - 1
        mVRAMEntries(i).BaseAddress = mVRAMEntries(i + 1).BaseAddress
        mVRAMEntries(i).Bank = mVRAMEntries(i + 1).Bank
        Set mVRAMEntries(i).Bitmap = mVRAMEntries(i + 1).Bitmap
        mVRAMEntries(i).Filename = mVRAMEntries(i + 1).Filename
    Next i
    
    ReDim Preserve mVRAMEntries(UBound(mVRAMEntries) - 1)

End Sub

Public Sub PackToBin(ByVal sFilename As String)

    On Error GoTo HandleErrors

    If sFilename = "" Then
        Exit Sub
    End If

    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Output As #nFilenum
    
    Dim i As Integer
    Dim j As Integer
    Dim strTab As String
    Dim strCount0 As String
    Dim strCount1 As String
    Dim c0 As String
    Dim c1 As String
    Dim strAddr As String
    Dim strWidth As String
    Dim strHeight As String
    
    strTab = Chr(vbKeyTab)
    
    For i = 1 To VRAMEntryCount
        If VRAMEntryBank(i) = 0 Then
            c0 = c0 & Format$(CStr(i), "00")
        Else
            c1 = c1 & Format$(CStr(i), "00")
        End If
    Next i
    
    Dim nFilesize As Integer
    nFilesize = 2 + (Len(c0) * 3) + (Len(c0) \ 2) + (Len(c1) * 3) + (Len(c1) \ 2)
    
    Print #nFilenum, ";Size:" & CStr(nFilesize)
    Print #nFilenum, ";********************************"
    Print #nFilenum, "; " & UCase$(GetTruncFilename(sFilename))
    Print #nFilenum, ";********************************"
    Print #nFilenum, ";" & strTab & "Author:" & strTab & "Mr. Yuk"
    Print #nFilenum, ";" & strTab & "(c)2000" & strTab & "Interactive Imagination"
    Print #nFilenum, ";" & strTab & "All rights reserved"
    Print #nFilenum, ""
    Print #nFilenum, ";********************************"
    
    Print #nFilenum, ""
    
    For i = 1 To VRAMEntryCount
        Print #nFilenum, strTab & "GLOBAL" & strTab & strTab & UCase$(Mid$(GetTruncFilename(VRAMEntryFilenames(i)), 1, Len(GetTruncFilename(VRAMEntryFilenames(i))) - 4))
    Next i
    
    Print #nFilenum, ""
    Print #nFilenum, UCase$(Mid$(GetTruncFilename(sFilename), 1, Len(GetTruncFilename(sFilename)) - 2))
    
    strCount0 = Hex(Len(c0) / 2)
    
    If Len(strCount0) < 2 Then
        strCount0 = "$0" & strCount0
    Else
        strCount0 = "$" & strCount0
    End If
    Print #nFilenum, strTab & "DB" & strTab & strTab & strCount0
    Print #nFilenum, ""
    
    For j = 1 To Len(c0) Step 2
        i = CInt(Mid$(c0, j, 2))
        Print #nFilenum, strTab & "DEFW" & strTab & "$" & Hex(VRAMEntryBaseAddress(i))
        Print #nFilenum, strTab & "DEFW" & strTab & "(" & UCase$(Mid$(GetTruncFilename(VRAMEntryFilenames(i)), 1, Len(GetTruncFilename(VRAMEntryFilenames(i))) - 4) & "&$FFFF)")
        
        Dim pnt As New clsPoint
        
        strWidth = Hex(VRAMEntryBitmap(i).width \ 8)
        
        If Len(strWidth) < 2 Then
            strWidth = "$0" & strWidth
        Else
            strWidth = "$" & strWidth
        End If
        Print #nFilenum, strTab & "DB" & strTab & strTab & strWidth
        
        strHeight = Hex(VRAMEntryBitmap(i).height \ 8)
        If Len(strHeight) < 2 Then
            strHeight = "$0" & strHeight
        Else
            strHeight = "$" & strHeight
        End If
        Print #nFilenum, strTab & "DB" & strTab & strTab & strHeight
        
        Print #nFilenum, strTab & "DB" & strTab & strTab & ":" & UCase$(Mid$(GetTruncFilename(VRAMEntryFilenames(i)), 1, Len(GetTruncFilename(VRAMEntryFilenames(i))) - 4))
        Print #nFilenum, ""
    Next j
    
    strCount1 = Hex(Len(c1) / 2)
    
    If Len(strCount1) < 2 Then
        strCount1 = "$0" & strCount1
    Else
        strCount1 = "$" & strCount1
    End If
    Print #nFilenum, strTab & "DB" & strTab & strTab & strCount1
    Print #nFilenum, ""
    
    For j = 1 To Len(c1) Step 2
        i = CInt(Mid$(c1, j, 2))
        Print #nFilenum, strTab & "DEFW" & strTab & "$" & Hex(VRAMEntryBaseAddress(i))
        Print #nFilenum, strTab & "DEFW" & strTab & "(" & UCase$(Mid$(GetTruncFilename(VRAMEntryFilenames(i)), 1, Len(GetTruncFilename(VRAMEntryFilenames(i))) - 4) & "&$FFFF)")
        
        strWidth = Hex(VRAMEntryBitmap(i).width \ 8)
        If Len(strWidth) < 2 Then
            strWidth = "$0" & strWidth
        Else
            strWidth = "$" & strWidth
        End If
        Print #nFilenum, strTab & "DB" & strTab & strTab & strWidth
        
        strHeight = Hex(VRAMEntryBitmap(i).height \ 8)
        If Len(strHeight) < 2 Then
            strHeight = "$0" & strHeight
        Else
            strHeight = "$" & strHeight
        End If
        Print #nFilenum, strTab & "DB" & strTab & strTab & strHeight
        
        Print #nFilenum, strTab & "DB" & strTab & strTab & ":" & UCase$(Mid$(GetTruncFilename(VRAMEntryFilenames(i)), 1, Len(GetTruncFilename(VRAMEntryFilenames(i))) - 4))
        Print #nFilenum, ""
    Next j
    
    Print #nFilenum, ";********************************"
    Print #nFilenum, strTab & "END"
    Print #nFilenum, ";********************************"
    
'name of vram file -ext
    
    'write bank0
        'DB total vram entries
            'DEFW address
            'DEFW ([name of bitmap file -extension]&$FFFF)
            'DB bitmap width in pixels in units of tiles
            'DB bitmap height in tiles
            'DB :[name of bitmap file -extension]
    
    'repeat for bank1
    
    'end
    
    Close #nFilenum

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBVRAM:PackToBin Error"
    Close #nFilenum
End Sub

Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient
    
End Sub

Public Property Let VRAMEntryBank(Index As Integer, iNewValue As Byte)

    mVRAMEntries(Index).Bank = iNewValue

End Property

Public Property Let VRAMEntryBaseAddress(Index As Integer, lNewValue As Long)

    mVRAMEntries(Index).BaseAddress = lNewValue

End Property

Public Property Set VRAMEntryBitmap(Index As Integer, oNewValue As clsGBBitmap)

    Set mVRAMEntries(Index).Bitmap = oNewValue

End Property

Public Property Let VRAMEntryFilenames(Index As Integer, sNewValue As String)

    mVRAMEntries(Index).Filename = sNewValue

End Property

Private Sub Class_Initialize()

    ReDim mVRAMEntries(0)

End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Sub intResource_Delete()

'***************************************************************************
'   Release memory when the class terminates
'***************************************************************************
    
    If Not mOffscreenBank0 Is Nothing Then
        mOffscreenBank0.Delete
    End If
    
    If Not mOffscreenBank1 Is Nothing Then
        mOffscreenBank1.Delete
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 384
        For j = 0 To 1
            If Not mBitmapFragments(i, j) Is Nothing Then
                mBitmapFragments(i, j).Delete
            End If
        Next j
    Next i

End Sub

Private Function intResource_Unpack(nFilenum As Integer) As Boolean
    
'***************************************************************************
'   Unpack the data from a .vrm file
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    intResource_Unpack = True
    
    Dim i As Integer
    Dim s As Integer
    Dim dummy As Integer
    
    Get #nFilenum, , dummy
    
    ReDim mVRAMEntries(dummy)
    
    For i = 1 To dummy
        Get #nFilenum, , s
        mVRAMEntries(i).Filename = String(s, " ")
        Get #nFilenum, , mVRAMEntries(i).Filename
        
        mVRAMEntries(i).Filename = msParentPath & "\Bitmaps\" & GetTruncFilename(mVRAMEntries(i).Filename)
        
        If gbBatchExport Then
            gbBatchExportVRAMLoad = True
        End If
        
        gbOpeningChild = True
        Set mVRAMEntries(i).Bitmap = gResourceCache.GetResourceFromFile(mVRAMEntries(i).Filename, Me)
        gbOpeningChild = False
        
        If gbBatchExport Then
            gbBatchExportVRAMLoad = False
        End If
        
        If Not gbBatchExport Then
            If mVRAMEntries(i).Bitmap Is Nothing Then
                intResource_Unpack = False
                Exit Function
            End If
        End If
        
        Get #nFilenum, , mVRAMEntries(i).BaseAddress
        Get #nFilenum, , mVRAMEntries(i).Bank
    Next i

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBVRAM:intResource_Unpack Error"
    intResource_Unpack = False
End Function

Private Sub intResource_Pack(nFilenum As Integer)

'***************************************************************************
'   Pack data into a .vrm file
'***************************************************************************
    
    Dim i As Integer
    Dim dummy As Integer
    
    Put #nFilenum, , CByte(GB_VRAM)
    
    dummy = CInt(UBound(mVRAMEntries))
    Put #nFilenum, , dummy
    
    For i = 1 To UBound(mVRAMEntries)
        dummy = CInt(Len(GetTruncFilename(mVRAMEntries(i).Filename)))
        Put #nFilenum, , dummy
        Put #nFilenum, , GetTruncFilename(mVRAMEntries(i).Filename)
        Put #nFilenum, , mVRAMEntries(i).BaseAddress
        Put #nFilenum, , mVRAMEntries(i).Bank
    Next i

End Sub

Private Property Set intResource_CacheObject(oCacheObject As clsResourceCacheObject)
    
    Set mResourceCacheObject = oCacheObject
    
End Property

Private Sub intResource_UpdateClients(iClient As intResourceClient)
    
'***************************************************************************
'   Update all of the resources clients
'***************************************************************************
    
    If Not mResourceCacheObject Is Nothing Then
        mResourceCacheObject.UpdateClients iClient
    End If
    
End Sub

Private Property Get intResource_ResourceType() As GB_FILETYPES
    
'***************************************************************************
'   Returns the resource type
'***************************************************************************
    
    intResource_ResourceType = GB_VRAM
    
End Property

Public Property Get VRAMEntryFilenames(Index As Integer) As String
    
    VRAMEntryFilenames = mVRAMEntries(Index).Filename
    
End Property

Public Property Get VRAMEntryBaseAddress(Index As Integer) As Long
    
    VRAMEntryBaseAddress = mVRAMEntries(Index).BaseAddress
    
End Property

Public Property Get VRAMEntryBank(Index As Integer) As Byte
    
    VRAMEntryBank = mVRAMEntries(Index).Bank
    
End Property

Public Property Get VRAMEntryCount() As Integer
    
    VRAMEntryCount = UBound(mVRAMEntries)
    
End Property

Public Property Get VRAMEntryBitmap(Index As Integer) As clsGBBitmap
    
    Set VRAMEntryBitmap = mVRAMEntries(Index).Bitmap
    
End Property

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)
    
    intResource_UpdateClients Nothing
    
End Sub

Public Function AddVRAMEntry(ByVal sFilename As String, BaseAddress As Long, Bank As Byte) As Boolean

'***************************************************************************
'   Add an entry into the VRAM
'***************************************************************************

    ReDim Preserve mVRAMEntries(UBound(mVRAMEntries) + 1)
    
    mVRAMEntries(UBound(mVRAMEntries)).Filename = sFilename
    mVRAMEntries(UBound(mVRAMEntries)).BaseAddress = BaseAddress
    mVRAMEntries(UBound(mVRAMEntries)).Bank = Bank
    Set mVRAMEntries(UBound(mVRAMEntries)).Bitmap = gResourceCache.GetResourceFromFile(mVRAMEntries(UBound(mVRAMEntries)).Filename, Me)
    
    If mVRAMEntries(UBound(mVRAMEntries)).Bitmap Is Nothing Then
        ReDim Preserve mVRAMEntries(UBound(mVRAMEntries) - 1)
        AddVRAMEntry = False
        Exit Function
    End If

    AddVRAMEntry = True

End Function
