VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBBackground"
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
    
    Private Type tTile
        BitmapFragmentIndex As Integer
        VRAMEntryBank As Byte
        VRAMEntryAddress As Long
        PaletteID As Byte
        XFlip As Byte
        YFlip As Byte
        Priority As Byte
    End Type

    Private mtTiles() As tTile
    Private mnWidth As Integer
    Private mnHeight As Integer
    Private mtBackgroundType As GB_BACKGROUNDTYPES
    Private msPaletteFile As String
    Private msVRAMFile As String
    Private mOffscreen As New clsOffscreen
    Private mMapOffscreen As New clsOffscreen
    Private mGBVRAM As New clsGBVRAM
    Private mGBPalette As New clsGBPalette
    Private msParentPath As String
    
'***************************************************************************
'   Resource cache object pointer
'***************************************************************************

    Private mResourceCacheObject As clsResourceCacheObject
    
Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property

Public Property Get nWidth() As Integer

    nWidth = mnWidth

End Property

Public Property Get nHeight() As Integer

    nHeight = mnHeight

End Property

Public Property Let nWidth(nNewValue As Integer)

    On Error GoTo HandleErrors

    Dim X As Integer
    Dim Y As Integer
    Dim w As Integer
    Dim h As Integer

    w = mnWidth
    h = mnHeight
    ReDim dummy(w, h) As tTile
    
    For Y = 0 To h
        For X = 0 To w
            dummy(X, Y) = mtTiles(X, Y)
        Next X
    Next Y
    
    mnWidth = nNewValue
    ReDim mtTiles(mnWidth, mnHeight)
    
    For Y = 0 To h
        For X = 0 To w
            mtTiles(X, Y) = dummy(X, Y)
        Next X
    Next Y

Exit Property

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBackground:nWidth Error"
End Property

Public Property Let nHeight(nNewValue As Integer)

    Dim X As Integer
    Dim Y As Integer
    Dim w As Integer
    Dim h As Integer

    w = mnWidth
    h = mnHeight
    ReDim dummy(w, h) As tTile
    
    For Y = 0 To h
        For X = 0 To w
            dummy(X, Y) = mtTiles(X, Y)
        Next X
    Next Y
    
    mnHeight = nNewValue
    ReDim mtTiles(mnWidth, mnHeight)
    
    For Y = 0 To h
        For X = 0 To w
            mtTiles(X, Y) = dummy(X, Y)
        Next X
    Next Y
    
End Property

Public Sub PackToBin(sFilename As String)

    If sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    On Error GoTo HandleErrors
    
    Dim px As Integer
    Dim py As Integer
    Dim tx As Integer
    Dim ty As Integer
    Dim code As Long
    
    If mtBackgroundType = GB_PATTERNBG Then
        
        Open sFilename For Binary Access Write As #nFilenum
    
        For py = 0 To 15
            For px = 0 To 15
                For ty = 0 To 1
                    For tx = 0 To 1
            
                        code = (((VRAMEntryAddress((px * 2) + tx + 1, (py * 2) + ty + 1) - 32768) / 16) - 128)
    
                        If code < 0 Then
                            code = 0
                        End If
    
                        If code < 128 Then
                            code = code + 128
                        Else
                            code = code - 128
                        End If
            
                        Put #nFilenum, , CByte(code)
                    
                    Next tx
                Next ty
            Next px
        Next py
    
        Dim patX As Integer
        Dim patY As Integer
        Dim str As String
    
        For py = 0 To 15
            For px = 0 To 15
                For ty = 0 To 1
                    For tx = 0 To 1
                    
                        patX = (px * 2) + tx + 1
                        patY = (py * 2) + ty + 1
                    
                        Dim spung As Byte
                    
                        spung = CByte(PaletteID(patX, patY))
                        spung = CByte(spung + ((VRAMEntryBank(patX, patY) * 8)))
                        spung = CByte(spung + (-((XFlip(patX, patY) <> 0) * 32)))
                        spung = CByte(spung + (-((YFlip(patX, patY) <> 0) * 64)))
                        spung = CByte(spung + ((Priority(patX, patY) * 128)))
                    
                        Put #nFilenum, , spung
                    
                    Next tx
                Next ty
            Next px
        Next py
    
        Close #nFilenum
    
    Else
     
        Dim strCode As String
        Dim strSpung As String
        Dim strTab As String
     
        Open sFilename For Output As #nFilenum
    
        strTab = Chr(vbKeyTab)
        
        'determine filesize
        Dim nFilesize As Integer
        nFilesize = (nHeight * nWidth) * 2
        
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
    
        Print #nFilenum, ";Width"
        Print #nFilenum, ";DB " & nWidth
        Print #nFilenum, ";Height"
        Print #nFilenum, ";DB " & nHeight
    
        Dim matt As String
        Dim matti As Integer
        matt = UCase$(GetTruncFilename(sFilename))
        
        For matti = Len(matt) To 1 Step -1
            If Mid$(matt, matti, 1) = "." Then
                matt = Mid$(matt, 1, matti - 1)
                Exit For
            End If
        Next matti
    
        Print #nFilenum, matt
    
        For py = 1 To nHeight
            Print #nFilenum, strTab & "DB" & strTab;
            For px = 1 To nWidth
            
                code = (((VRAMEntryAddress(px, py) - 32768) / 16) - 128)
    
                If code < 0 Then
                    code = 0
                End If
    
                If code < 128 Then
                    code = code + 128
                Else
                    code = code - 128
                End If
    
                strCode = Hex(code)
                If Len(strCode) < 2 Then
                    strCode = "$0" & strCode
                Else
                    strCode = "$" & strCode
                End If
    
                If px = nWidth Then
                    Print #nFilenum, strCode
                Else
                    Print #nFilenum, strCode & ",";
                End If
                    
            Next px
        Next py
    
        Print #nFilenum, ""
    
        Print #nFilenum, matt & "_ATTRIB"
    
        For patY = 1 To nHeight
            Print #nFilenum, strTab & "DB" & strTab;
            For patX = 1 To nWidth
                    
                spung = PaletteID(patX, patY)
                spung = CByte(spung + VRAMEntryBank(patX, patY) * 8)
                spung = CByte(spung + -(XFlip(patX, patY) <> 0) * 32)
                spung = CByte(spung + -(YFlip(patX, patY) <> 0) * 64)
                spung = CByte(spung + Priority(patX, patY) * 128)
            
                strSpung = Hex(spung)
                If Len(strSpung) < 2 Then
                    strSpung = "$0" & strSpung
                Else
                    strSpung = "$" & strSpung
                End If
                
                If patX = nWidth Then
                    Print #nFilenum, strSpung
                Else
                    Print #nFilenum, strSpung & ",";
                End If
                    
            Next patX
        Next patY
    
        Print #nFilenum, ""
        Print #nFilenum, ";********************************"
    
        Print #nFilenum, strTab & "END"
        Print #nFilenum, ";********************************"

        Close #nFilenum
    End If
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBackground:PackToBin Error"
End Sub


Public Property Get GBPalette() As clsGBPalette

    Set GBPalette = mGBPalette

End Property

Public Property Set GBPalette(oNewValue As clsGBPalette)

    Set mGBPalette = oNewValue

End Property

Public Property Get GBVRAM() As clsGBVRAM

    Set GBVRAM = mGBVRAM

End Property

Public Property Set GBVRAM(oNewValue As clsGBVRAM)

    Set mGBVRAM = oNewValue

End Property

Public Property Get MapOffscreen() As clsOffscreen

    Set MapOffscreen = mMapOffscreen

End Property

Public Property Get Offscreen() As clsOffscreen

    Set Offscreen = mOffscreen

End Property

Public Property Set Offscreen(oNewValue As clsOffscreen)

    Set mOffscreen = oNewValue

End Property

Public Sub PackToBinRLE(ByVal sFilename As String)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim px As Integer
    Dim py As Integer
    Dim patX As Integer
    Dim patY As Integer
    Dim nFilenum As Integer
    
    nFilenum = FreeFile

    Open App.Path & "\TEMPBG2.bin" For Binary As #nFilenum

        For py = 1 To nHeight
            For px = 1 To nWidth
            
                Dim code As Integer
            
                code = (((VRAMEntryAddress(px, py) - 32768) / 16) - 128)
    
                If code < 0 Then
                    code = 0
                End If
    
                If code < 128 Then
                    code = code + 128
                Else
                    code = code - 128
                End If
    
                Put #nFilenum, , CByte(code)
                
            Next px
        Next py
    
    Close #nFilenum

    nFilenum = FreeFile

    Open App.Path & "\TEMPBG1.bin" For Binary As #nFilenum

        For patY = 1 To nHeight
            For patX = 1 To nWidth
                
                Dim spung As Byte
                
                spung = PaletteID(patX, patY)
                spung = CByte(spung + VRAMEntryBank(patX, patY) * 8)
                spung = CByte(spung + -(XFlip(patX, patY) <> 0) * 32)
                spung = CByte(spung + -(YFlip(patX, patY) <> 0) * 64)
                spung = CByte(spung + Priority(patX, patY) * 128)
            
                Put #nFilenum, , spung
                    
            Next patX
        Next patY

    Close #nFilenum

    PackRLE App.Path & "\TEMPBG1.bin", App.Path & "\TEMPBGRLE1.bin", 0
    PackRLE App.Path & "\TEMPBG2.bin", App.Path & "\TEMPBGRLE2.bin", 0
    
'Combine the two RLE files

    nFilenum = FreeFile
    
    Open App.Path & "\TEMPBGRLE1.bin" For Binary As #nFilenum
    
        Dim dummy As Integer
        dummy = LOF(nFilenum)
        
        ReDim iFile1(LOF(nFilenum)) As Byte
        
        For i = 0 To (dummy - 1)
            Get #nFilenum, , iFile1(i)
        Next i
    
    Close #nFilenum
    
    nFilenum = FreeFile
    
    Open App.Path & "\TEMPBGRLE2.bin" For Binary As #nFilenum
    
        dummy = LOF(nFilenum)
        
        ReDim iFile2(dummy) As Byte
        
        For i = 0 To (dummy - 1)
            Get #nFilenum, , iFile2(i)
        Next i
    
    Close #nFilenum
    
    For i = Len(sFilename) To 1 Step -1
        If Mid$(sFilename, i, 1) = "." Then
            sFilename = Mid$(sFilename, 1, i - 1) & ".bin"
            Exit For
        End If
    Next i
    
    Open sFilename For Binary As #nFilenum
    
        For i = 0 To (UBound(iFile1) - 1)
            Put #nFilenum, , iFile1(i)
        Next i
    
        For i = 0 To (UBound(iFile2) - 1)
            Put #nFilenum, , iFile2(i)
        Next i

rExit:
    
    Close #nFilenum
    
    DeleteFile App.Path & "\TEMPBG1.bin"
    DeleteFile App.Path & "\TEMPBG2.bin"
    DeleteFile App.Path & "\TEMPBGRLE1.bin"
    DeleteFile App.Path & "\TEMPBGRLE2.bin"

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBBackground:PackToBinRLE Error"
    GoTo rExit
End Sub

Public Property Get sPaletteFile() As String

    sPaletteFile = msPaletteFile

End Property

Public Property Let sPaletteFile(sNewValue As String)

    msPaletteFile = sNewValue

End Property

Public Property Get sVRAMFile() As String

    sVRAMFile = msVRAMFile

End Property

Public Property Let sVRAMFile(sNewValue As String)

    msVRAMFile = sNewValue

End Property

    
Public Property Let BitmapFragmentIndex(X As Integer, Y As Integer, nNewValue As Integer)

    mtTiles(X, Y).BitmapFragmentIndex = nNewValue

End Property

Public Property Get BitmapFragmentIndex(X As Integer, Y As Integer) As Integer

    BitmapFragmentIndex = mtTiles(X, Y).BitmapFragmentIndex

End Property

Public Property Let tBackgroundType(tNewValue As GB_BACKGROUNDTYPES)

    mtBackgroundType = tNewValue

End Property

Public Property Get tBackgroundType() As GB_BACKGROUNDTYPES

    tBackgroundType = mtBackgroundType

End Property

Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient

End Sub

Public Property Get VRAMEntryAddress(PatternX As Integer, PatternY As Integer) As Long

    VRAMEntryAddress = mtTiles(PatternX, PatternY).VRAMEntryAddress

End Property

Public Property Get PaletteID(PatternX As Integer, PatternY As Integer) As Long

    PaletteID = mtTiles(PatternX, PatternY).PaletteID

End Property

Public Property Get VRAMEntryBank(PatternX As Integer, PatternY As Integer) As Byte

    VRAMEntryBank = mtTiles(PatternX, PatternY).VRAMEntryBank

End Property

Public Property Get XFlip(PatternX As Integer, PatternY As Integer) As Byte

    XFlip = mtTiles(PatternX, PatternY).XFlip

End Property

Public Property Get Priority(PatternX As Integer, PatternY As Integer) As Byte

    Priority = mtTiles(PatternX, PatternY).Priority

End Property

Public Property Get YFlip(PatternX As Integer, PatternY As Integer) As Byte

    YFlip = mtTiles(PatternX, PatternY).YFlip

End Property

Public Property Let VRAMEntryAddress(PatternX As Integer, PatternY As Integer, nNewValue As Long)

    mtTiles(PatternX, PatternY).VRAMEntryAddress = nNewValue

End Property

Public Property Let PaletteID(PatternX As Integer, PatternY As Integer, nNewValue As Long)

    mtTiles(PatternX, PatternY).PaletteID = nNewValue

End Property

Public Property Let VRAMEntryBank(PatternX As Integer, PatternY As Integer, nNewValue As Byte)

    mtTiles(PatternX, PatternY).VRAMEntryBank = nNewValue

End Property

Public Property Let XFlip(PatternX As Integer, PatternY As Integer, nNewValue As Byte)

    mtTiles(PatternX, PatternY).XFlip = nNewValue

End Property

Public Property Let YFlip(PatternX As Integer, PatternY As Integer, nNewValue As Byte)

    mtTiles(PatternX, PatternY).YFlip = nNewValue

End Property

Public Property Let Priority(PatternX As Integer, PatternY As Integer, nNewValue As Byte)

    mtTiles(PatternX, PatternY).Priority = nNewValue

End Property

Private Sub Class_Initialize()

    ReDim mtTiles(0, 0)

End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Sub intResource_Delete()
    
'***************************************************************************
'   Release memory when the class is terminated
'***************************************************************************
    
    If Not mOffscreen Is Nothing Then
        mOffscreen.Delete
    End If
    
    If Not mMapOffscreen Is Nothing Then
        mMapOffscreen.Delete
    End If
    
End Sub

Public Sub Delete()

    intResource_Delete

End Sub

Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Private Function intResource_Unpack(nFilenum As Integer) As Boolean
    
'***************************************************************************
'   Unpack the background from a .pat or .bg file
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    intResource_Unpack = True
    
    Dim X As Integer
    Dim Y As Integer
    Dim j As Integer
    Dim s As Integer
    Dim d As Integer
    Dim dVRAM As clsGBVRAM
    Dim sErrMsg As String
    Dim sDir As String
    Dim bOldFormat As Boolean
    
    Get #nFilenum, , s
    If s = 1 Then
        bOldFormat = False
        Get #nFilenum, , s
    Else
        bOldFormat = True
    End If
    
'Get VRAM filename
    msVRAMFile = String(s, " ")
    Get #nFilenum, , msVRAMFile
    
    msVRAMFile = msParentPath & "\VRAMs\" & GetTruncFilename(msVRAMFile)
    
    If Not gbBatchExport Then
        If msVRAMFile <> "" Then
            sErrMsg = msVRAMFile
            
            sDir = "VRAMs"
            gbOpeningChild = True
            Set mGBVRAM = gResourceCache.GetResourceFromFile(msVRAMFile, Me)
            gbOpeningChild = False
            
            If mGBVRAM.CacheObject Is Nothing Then
                intResource_Unpack = False
                Exit Function
            End If
                        
        End If
    End If

    msVRAMFile = GetTruncFilename(msVRAMFile)

'Get Palette filename
    Get #nFilenum, , s
    msPaletteFile = String(s, " ")
    Get #nFilenum, , msPaletteFile
    
    msPaletteFile = msParentPath & "\Palettes\" & GetTruncFilename(msPaletteFile)
    
    If Not gbBatchExport Then
        If msPaletteFile <> "" Then
            sErrMsg = msPaletteFile
            
            sDir = "Palettes"
            gbOpeningChild = True
            Set mGBPalette = gResourceCache.GetResourceFromFile(msPaletteFile, Me)
            gbOpeningChild = False
            
            If mGBPalette.CacheObject Is Nothing Then
                intResource_Unpack = False
                Exit Function
            End If
            
        End If
    End If

    msPaletteFile = GetTruncFilename(msPaletteFile)

'Get dimensions
    Get #nFilenum, , d
    nWidth = d
    
    Get #nFilenum, , d
    nHeight = d

'Get data for each tile on the background
    For Y = 1 To mnHeight
        For X = 1 To mnWidth
            Get #nFilenum, , mtTiles(X, Y).XFlip
            Get #nFilenum, , mtTiles(X, Y).YFlip
            Get #nFilenum, , mtTiles(X, Y).PaletteID
            Get #nFilenum, , mtTiles(X, Y).VRAMEntryAddress
            Get #nFilenum, , mtTiles(X, Y).VRAMEntryBank
            
            If Not bOldFormat Then
                Get #nFilenum, , mtTiles(X, Y).Priority
            End If
            
        Next X
    Next Y

Exit Function

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sErrMsg & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = msParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                intResource_Unpack = False
                Exit Function
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, msParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, msParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            If sDir = "VRAMs" Then
                msVRAMFile = .Filename
            ElseIf sDir = "Palettes" Then
                msPaletteFile = .Filename
            End If
                        
            Resume
            
        End With
    
    Else
        MsgBox Err.Description, vbCritical, "clsGBBackground:intResourceClient_Update Error"
        intResource_Unpack = False
    End If
End Function

Private Sub intResource_Pack(nFilenum As Integer)

'***************************************************************************
'   Pack resource to a .pat or .bg file
'***************************************************************************

    Dim X As Integer
    Dim Y As Integer
    Dim s As Integer
    Dim j As Integer

    If tBackgroundType = GB_PATTERNBG Then
        Put #nFilenum, , CByte(GB_PATTERN)
    Else
        Put #nFilenum, , CByte(GB_BG)
    End If

    s = 1
    Put #nFilenum, , s

    s = Len(GetTruncFilename(msVRAMFile))
    Put #nFilenum, , s
    Put #nFilenum, , UCase$(GetTruncFilename(msVRAMFile))
    
    s = Len(GetTruncFilename(msPaletteFile))
    Put #nFilenum, , s
    Put #nFilenum, , UCase$(GetTruncFilename(msPaletteFile))
    
    Put #nFilenum, , CInt(mnWidth)
    Put #nFilenum, , CInt(mnHeight)
    
    For Y = 1 To mnHeight
        For X = 1 To mnWidth
            Put #nFilenum, , mtTiles(X, Y).XFlip
            Put #nFilenum, , mtTiles(X, Y).YFlip
            Put #nFilenum, , mtTiles(X, Y).PaletteID
            Put #nFilenum, , mtTiles(X, Y).VRAMEntryAddress
            Put #nFilenum, , mtTiles(X, Y).VRAMEntryBank
            Put #nFilenum, , mtTiles(X, Y).Priority
        Next X
    Next Y
    
End Sub

Private Property Set intResource_CacheObject(oNewValue As clsResourceCacheObject)
    
    Set mResourceCacheObject = oNewValue
    
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
'   Return the resource type
'***************************************************************************
    
    If tBackgroundType = GB_RAWBG Then
        intResource_ResourceType = GB_BG
    ElseIf tBackgroundType = GB_PATTERNBG Then
        intResource_ResourceType = GB_PATTERN
    End If
    
End Property

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)
    
'***************************************************************************
'   Update clients of the resource
'***************************************************************************
    
    intResource_UpdateClients Nothing

End Sub
