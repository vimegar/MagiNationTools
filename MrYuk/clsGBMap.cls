VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBMap"
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
    
    Private miMapData() As Integer
    Private miUndoData() As Integer
    Private miRedoData() As Integer
    Private mbRedoFlag As Boolean
    Private mnWidth As Integer
    Private mnHeight As Integer
    Private msPatternFile As String

    Private mGBBackground As New clsGBBackground
    Private mOffscreen As New clsOffscreen
    Private mbUndoFlag As Boolean
    Private msParentPath As String

'***************************************************************************
'   Resource cache object pointer
'***************************************************************************

    Private mResourceCacheObject As clsResourceCacheObject



Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient

End Sub

Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Public Sub Delete()

    intResource_Delete

End Sub


Public Property Get MapData(Index As Integer) As Integer

    MapData = miMapData(Index)

End Property

Public Property Let MapData(Index As Integer, nNewValue As Integer)

    miMapData(Index) = nNewValue
    
End Property

Public Property Set GBBackground(oNewValue As clsGBBackground)

    Set mGBBackground = oNewValue

End Property

Public Property Get GBBackground() As clsGBBackground

    Set GBBackground = mGBBackground

End Property



Public Property Set Offscreen(oNewValue As clsOffscreen)

    Set mOffscreen = oNewValue

End Property

Public Property Get Offscreen() As clsOffscreen

    Set Offscreen = mOffscreen

End Property

Public Sub PackToBin(ByVal sFilename As String)

'u8  Scroll_WriteLib ( ScrollPtr scroll, char* fileName )
'    {
'    BinFile     binFile;
'    u16         x, y;

'    if ( !BinFile_Create ( &binFile, fileName )) return 0;

'    for ( y = 0; y < scroll->YTiles; y++ )
'        for ( x = 0; x < scroll->XTiles; x++ )
'            BinFile_Write ( &binFile, &scroll->Tiles[x][y], sizeof ( u8 ));

'    BinFile_Close ( &binFile );

'    return 1;
'    }

    On Error GoTo HandleErrors

    If sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Binary Access Write As #nFilenum
    
    Dim X As Integer
    Dim Y As Integer
    
    Put #nFilenum, , CByte(mnWidth)
    Put #nFilenum, , CInt(mnWidth * 9)
    
    Dim b As Byte
    
    If mnWidth - 11 < 0 Then
        b = 255
    Else
        b = mnWidth - 11
    End If
    
    Put #nFilenum, , CByte(b)
    
    If mnHeight - 10 < 0 Then
        b = 255
    Else
        b = mnHeight - 10
    End If
    
    Put #nFilenum, , CByte(b)
    Put #nFilenum, , CInt(mnWidth * mnHeight)
   
    For Y = 0 To mnHeight - 1
        For X = 0 To mnWidth - 1
            Put #nFilenum, , CByte(miMapData(X + (Y * mnWidth) + 1))
        Next X
    Next Y

    Close #nFilenum

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBMap:PackToBin Error"
End Sub

Public Sub RestoreUndoState()

    Dim i As Integer
    
    If Not mbRedoFlag Then
        ReDim miRedoData(UBound(miUndoData))
        
        For i = 0 To UBound(miRedoData)
            miRedoData(i) = miMapData(i)
        Next i
        
        For i = 0 To UBound(miUndoData)
            miMapData(i) = miUndoData(i)
        Next i
        
        mbRedoFlag = True
    Else
        
        For i = 0 To UBound(miRedoData)
            miMapData(i) = miRedoData(i)
        Next i
        
        mbRedoFlag = False
    End If
    

End Sub

Public Sub SaveUndoState()

    If Not mbUndoFlag Then
        Exit Sub
    End If
    
    ReDim miUndoData(UBound(miMapData))
    
    Dim i As Integer
    
    For i = 0 To UBound(miMapData)
        miUndoData(i) = miMapData(i)
    Next i
    
End Sub

Public Property Let sPatternFile(sNewValue As String)

    msPatternFile = sNewValue

End Property

Public Property Get sPatternFile() As String

    sPatternFile = msPatternFile

End Property

Public Property Let UndoFlag(bNewValue As Boolean)

    mbUndoFlag = bNewValue
    mbRedoFlag = Not mbUndoFlag

End Property

Public Property Get width() As Integer

    width = mnWidth
    
End Property

Public Property Let width(nNewValue As Integer)

    mnWidth = nNewValue

    ReDim Preserve miMapData(mnWidth * mnHeight)
    
    Dim X As Integer
    Dim Y As Integer
    
    For Y = 0 To mnHeight - 1
        For X = 0 To mnWidth - 1
            miMapData(X + (Y * mnWidth) + 1) = 0
        Next X
    Next Y

End Property

Public Property Get height() As Integer

    height = mnHeight

End Property

Public Property Let height(nNewValue As Integer)

    mnHeight = nNewValue

    ReDim Preserve miMapData(mnWidth * mnHeight)

    Dim X As Integer
    Dim Y As Integer
    
    For Y = 0 To mnHeight - 1
        For X = 0 To mnWidth - 1
            miMapData(X + (Y * mnWidth) + 1) = 0
        Next X
    Next Y

End Property

Private Sub Class_Initialize()

'***************************************************************************
'   Initialize the map data
'***************************************************************************
    
    ReDim miMapData(0)
    ReDim miUndoData(0)
    ReDim miRedoData(0)

End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Sub intResource_Delete()

'***************************************************************************
'   Release memory when the class terminates
'***************************************************************************

    If Not mOffscreen Is Nothing Then
        mOffscreen.Delete
    End If

End Sub

Private Function intResource_Unpack(nFilenum As Integer) As Boolean
    
'***************************************************************************
'   Unpack the map resource from a .map file
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    intResource_Unpack = True
    
    Dim X As Integer
    Dim Y As Integer
    Dim s As Integer
    Dim bOldFormat As Boolean
    
    Get #nFilenum, , s
    
    If s = 1 Then
        bOldFormat = False
        Get #nFilenum, , s
    Else
        bOldFormat = True
    End If
    
    msPatternFile = String(s, " ")
    Get #nFilenum, , msPatternFile
    
    msPatternFile = msParentPath & "\Patterns\" & GetTruncFilename(msPatternFile)
    
    If Not gbBatchExport Then
        If msPatternFile <> "" Then
            
            gbOpeningChild = True
            Set mGBBackground = gResourceCache.GetResourceFromFile(msPatternFile, Me)
            gbOpeningChild = False
            
            If mGBBackground.GBVRAM.CacheObject Is Nothing Or mGBBackground.GBPalette.CacheObject Is Nothing Then
                intResource_Unpack = False
                Exit Function
            End If
            
            mGBBackground.tBackgroundType = GB_PATTERNBG
        End If
    End If
    
    Get #nFilenum, , mnWidth
    Get #nFilenum, , mnHeight
    
    ReDim miMapData(mnWidth * mnHeight)
    
    For Y = 1 To mnHeight
        For X = 1 To mnWidth
            Dim i As Byte
            Get #nFilenum, , i
            miMapData(X + ((Y - 1) * mnWidth)) = i + (bOldFormat = True)
        Next X
    Next Y
    
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBMap:intResource_Unpack Error"
    intResource_Unpack = False
End Function

Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property

Private Sub intResource_Pack(nFilenum As Integer)

'***************************************************************************
'   Pack the map resource into a .map file
'***************************************************************************

    On Error GoTo HandleErrors

    Dim X As Integer
    Dim Y As Integer
    Dim s As Integer
    
    Put #nFilenum, , CByte(GB_MAP)
    
    s = 1
    Put #nFilenum, , s
    
    s = Len(msPatternFile)
    Put #nFilenum, , s
    Put #nFilenum, , msPatternFile
    
    Put #nFilenum, , mnWidth
    Put #nFilenum, , mnHeight
    
    For Y = 1 To mnHeight
        For X = 1 To mnWidth
            Put #nFilenum, , CByte(miMapData(X + ((Y - 1) * mnWidth)))
        Next X
    Next Y
    
Exit Sub
    
HandleErrors:
    If Err.Description = "Overflow" Then
        Put #nFilenum, , CByte(1)
        Resume Next
    End If
    MsgBox Err.Description, vbCritical, "clsGBMap:intResource_Pack Error"
End Sub

Private Property Set intResource_CacheObject(oCacheObject As clsResourceCacheObject)
    
    Set mResourceCacheObject = oCacheObject
    
End Property

Private Sub intResource_UpdateClients(iClient As intResourceClient)
    
'***************************************************************************
'   Update all the clients of the current resource
'***************************************************************************
    
    mResourceCacheObject.UpdateClients iClient
    
End Sub

Private Property Get intResource_ResourceType() As GB_FILETYPES
    
'***************************************************************************
'   Returns the type of resource this is
'***************************************************************************
    
    intResource_ResourceType = GB_MAP
    
End Property

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    intResource_UpdateClients Nothing
    
End Sub
