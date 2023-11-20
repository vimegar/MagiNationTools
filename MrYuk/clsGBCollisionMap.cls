VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBCollisionMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements intResource
Implements intResourceClient

Private mGBMap As New clsGBMap
Private miCollisionData() As Byte
Private miUndoData() As Byte
Private miRedoData() As Byte
Private msMapFile As String
Private mbUndoFlag As Boolean
Private mbRedoFlag As Boolean
Private msParentPath As String

Private mGBCollisionCodes As New clsGBCollisionCodes
Private msCollisionCodesFile As String

Private mResourceCacheObject As clsResourceCacheObject


Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Public Property Let UndoFlag(bNewValue As Boolean)

    mbUndoFlag = bNewValue
    mbRedoFlag = Not mbUndoFlag

End Property
Public Sub RestoreUndoState()

    Dim i As Integer
    
    If Not mbRedoFlag Then
        ReDim miRedoData(UBound(miUndoData))
        
        For i = 0 To UBound(miRedoData)
            miRedoData(i) = miCollisionData(i)
        Next i
        
        For i = 0 To UBound(miUndoData)
            miCollisionData(i) = miUndoData(i)
        Next i
        
        mbRedoFlag = True
    Else
        
        For i = 0 To UBound(miRedoData)
            miCollisionData(i) = miRedoData(i)
        Next i
        
        mbRedoFlag = False
    End If

End Sub
Public Sub SaveUndoState()

    If Not mbUndoFlag Then
        Exit Sub
    End If
    
    ReDim miUndoData(UBound(miCollisionData))
    
    Dim i As Integer
    
    For i = 0 To UBound(miCollisionData)
        miUndoData(i) = miCollisionData(i)
    Next i
    
End Sub
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

    If sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Binary Access Write As #nFilenum
    
    Dim X As Integer
    Dim Y As Integer
    Dim d As Integer
    
    Put #nFilenum, , CByte(mGBMap.width)
    Put #nFilenum, , CInt(mGBMap.width * mGBMap.height)
   
    For Y = 0 To mGBMap.height - 1
        For X = 0 To mGBMap.width - 1
            d = CollisionData(X + (Y * mGBMap.width) + 1)
            If d >= 0 Then
                Put #nFilenum, , CByte(d)
            End If
        Next X
    Next Y

    Close #nFilenum
    
End Sub

Public Property Get GBCollisionCodes() As clsGBCollisionCodes

    Set GBCollisionCodes = mGBCollisionCodes

End Property

Public Property Set GBCollisionCodes(oNewValue As clsGBCollisionCodes)

    Set mGBCollisionCodes = oNewValue

End Property

Public Property Get sCollisionCodesFile() As String

    sCollisionCodesFile = msCollisionCodesFile

End Property

Public Property Let sCollisionCodesFile(sNewValue As String)

    msCollisionCodesFile = sNewValue

End Property

Public Property Let CollisionData(Index As Integer, iNewValue As Byte)

    miCollisionData(Index) = iNewValue

End Property

Public Property Get CollisionData(Index As Integer) As Byte

    CollisionData = miCollisionData(Index)

End Property

Public Sub Delete()

    intResource_Delete

End Sub

Public Property Set GBMap(oNewValue As clsGBMap)

    Set mGBMap = oNewValue
    
    ReDim Preserve miCollisionData(mGBMap.width * mGBMap.height)

End Property

Public Property Get GBMap() As clsGBMap

    Set GBMap = mGBMap

End Property

Public Property Let sMapFile(sNewValue As String)

    msMapFile = sNewValue

End Property

Public Property Get sMapFile() As String

    sMapFile = msMapFile

End Property

Private Sub Class_Initialize()

    ReDim miUndoData(0)
    ReDim miRedoData(0)

End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Property Set intResource_CacheObject(RHS As clsResourceCacheObject)

    Set mResourceCacheObject = RHS

End Property

Private Sub intResource_Delete()

    'gResourceCache.ReleaseClient Me
    
End Sub


Private Sub intResource_Pack(nFilenum As Integer)

    Dim i As Integer
    Dim s As Integer
    Dim d As Integer
    
    Put #nFilenum, , CByte(GB_COLLISIONMAP)
    
    s = Len(msMapFile)
    Put #nFilenum, , s
    Put #nFilenum, , msMapFile
    
    d = UBound(miCollisionData)
    Put #nFilenum, , d
    
    For i = 1 To d
        Put #nFilenum, , CByte(miCollisionData(i))
    Next i

    's = Len(msCollisionCodesFile)
    'Put #nFilenum, , s
    'Put #nFilenum, , msCollisionCodesFile

End Sub


Private Property Get intResource_ResourceType() As GB_FILETYPES

    intResource_ResourceType = GB_COLLISIONMAP
    
End Property


Private Function intResource_Unpack(nFilenum As Integer) As Boolean

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim s As Integer
    Dim d As Integer
    
    intResource_Unpack = True
    
    Get #nFilenum, , s
    msMapFile = String(s, " ")
    Get #nFilenum, , msMapFile
    
    msMapFile = msParentPath & "\Maps\" & GetTruncFilename(msMapFile)
    
    If gbBatchExport Then
        gbBatchExportColLoad = True
    End If
    
    gbOpeningChild = True
    Set mGBMap = gResourceCache.GetResourceFromFile(msMapFile, Me)
    gbOpeningChild = False
    
    If gbBatchExport Then
        gbBatchExportColLoad = False
    End If
        
    If Not gbBatchExport Then
        If mGBMap.CacheObject Is Nothing Then
            intResource_Unpack = False
            Exit Function
        End If
    End If
            
    Get #nFilenum, , d
    ReDim miCollisionData(d)
    
    For i = 1 To d
        Get #nFilenum, , miCollisionData(i)
    Next i

    Set mGBCollisionCodes = gGBCollisionCodes

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBCollisionMap:intResource_Unpack Error"
    intResource_Unpack = False
End Function

Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property

Private Sub intResource_UpdateClients(iClient As intResourceClient)

'***************************************************************************
'   Update all the clients of the current resource
'***************************************************************************
    
    mResourceCacheObject.UpdateClients iClient

End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    intResource_UpdateClients Nothing
    
End Sub


