VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResourceCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCacheObjectList As New Collection

Public Sub PopulateResLists()

    On Error GoTo HandleErrors

    Dim i As Integer
    
    With mdiMain
        .lstResources.Clear
        For i = 1 To CacheObjectList.count
            .lstResources.AddItem GetTruncFilename(mCacheObjectList(i).sFilename) & " (" & CStr(mCacheObjectList(i).ClientList.count) & ")"
        Next i
    End With

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsResourceCache:PopulateResLists Error"
End Sub

Public Property Get CacheObjectList() As Collection
    
    Set CacheObjectList = mCacheObjectList

End Property

Public Sub CloseCache()

    Dim i As Integer
    
    For i = 1 To CacheObjectList.count
        If i > CacheObjectList.count Then Exit For
        CacheObjectList(i).Delete
        CacheObjectList.remove i
        If i <= CacheObjectList.count Then
            i = i - 1
        End If
    Next i

    PopulateResLists

End Sub

Public Function GetResourceFromFile(sFilename As String, iClient As intResourceClient) As intResource

'***************************************************************************
'   Add a resource to the cache from a file
'***************************************************************************

    On Error GoTo HandleErrors

    Dim tempCursor As Long
    tempCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass

'Check for existing cache object before
    Dim i As Integer
    Dim str As String
    
    str = GetTruncFilename(sFilename)
    
    For i = 1 To TotalCacheObjects
        If (sFilename = mCacheObjectList(i).sFilename) Or (str = mCacheObjectList(i).sFilename) Then
            mCacheObjectList(i).AddClient iClient
            Set GetResourceFromFile = mCacheObjectList(i).Data
            GoTo rEnd
        End If
    Next i

'Unpacks the resource
    Dim NewResource As intResource
    Dim NewCacheObj As clsResourceCacheObject
    
    Set NewResource = UnpackFile(sFilename)
    
    str = GetTruncFilename(sFilename)
    
    If NewResource Is Nothing Then
        Set GetResourceFromFile = Nothing
        GoTo rEnd
    End If

'Check for existing cache object after
    For i = 1 To TotalCacheObjects
        If (sFilename = mCacheObjectList(i).sFilename) Or (str = mCacheObjectList(i).sFilename) Then
            mCacheObjectList(i).AddClient iClient
            Set GetResourceFromFile = mCacheObjectList(i).Data
            GoTo rEnd
        End If
    Next i
        
'Adds cache object to list of cache objects
    Set NewCacheObj = New clsResourceCacheObject
    mCacheObjectList.Add NewCacheObj
    Set NewResource.CacheObject = NewCacheObj
    
    Set NewCacheObj.Data = NewResource
    NewCacheObj.sFilename = str
    NewCacheObj.AddClient iClient
    Set GetResourceFromFile = NewResource
    
rEnd:
    PopulateResLists
    Screen.MousePointer = tempCursor
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsResourceCache:GetResourceFromFile Error"
    GoTo rEnd
End Function

Public Sub ReleaseClient(iClient As intResourceClient)

'***************************************************************************
'   Release a client from its ties to a resource
'***************************************************************************
    
    On Error GoTo HandleErrors

    Dim clientCursor As intResourceClient
    Dim cacheObj As clsResourceCacheObject

remove:
    
    For Each cacheObj In CacheObjectList
        For Each clientCursor In cacheObj.ClientList
            
            If clientCursor Is iClient Then
                
                cacheObj.ClientList.remove CollectionFind(cacheObj.ClientList, clientCursor)
                
                If cacheObj.ClientList.count <= 0 Then
                    ReleaseClient cacheObj.Data
                    cacheObj.Delete
                    GoTo remove
                End If
                
            End If
            
        Next clientCursor
    Next cacheObj

    Purge

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsResourceCache:ReleaseClient Error"
End Sub

Public Sub AddResourceToCache(ByVal sFilename As String, iResource As intResource, iClient As intResourceClient)

    On Error GoTo HandleErrors

'***************************************************************************
'   Add a specified resource to the cache
'***************************************************************************
    
    Dim tempCursor As Long
    tempCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    Dim addNewCacheObj As Boolean
    Dim indexNew As Integer

    addNewCacheObj = True
    
'Check for existing cache object
    Dim i As Integer
    Dim str As String
    
    str = GetTruncFilename(sFilename)
    
    For i = 1 To TotalCacheObjects
        If (sFilename = mCacheObjectList(i).sFilename) Or (str = mCacheObjectList(i).sFilename) Then
            addNewCacheObj = False
            indexNew = i
        End If
    Next i
    
    If addNewCacheObj = True Then
        mCacheObjectList.Add New clsResourceCacheObject
        indexNew = mCacheObjectList.count
    End If
        
    Set mCacheObjectList(indexNew).Data = iResource
    mCacheObjectList(indexNew).sFilename = sFilename
    Set iResource.CacheObject = mCacheObjectList(indexNew)
    
'Add client to client list
    mCacheObjectList(indexNew).AddClient iClient
    
    iResource.UpdateClients iClient
    
    PopulateResLists
    Screen.MousePointer = tempCursor

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsResourceCache:AddResourceToCache Error"
End Sub

Public Sub Purge()
    
'***************************************************************************
'   Delete all resource that do not belong in the cache
'***************************************************************************
    
    Dim i As Integer
    
    For i = 1 To CacheObjectList.count
        If i > CacheObjectList.count Then Exit For
        If CacheObjectList(i).ClientList.count = 0 Then
            
            CacheObjectList.remove i
            
            If i <= CacheObjectList.count Then
                i = i - 1
            End If
            
        End If
    Next i
    
    PopulateResLists
    
End Sub

Public Property Get TotalCacheObjects() As Integer

    TotalCacheObjects = mCacheObjectList.count
    
End Property
