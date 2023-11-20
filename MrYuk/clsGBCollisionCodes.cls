VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBCollisionCodes"
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

    Private mBitmap As New StdPicture
    Private msBitmapFile As String

'***************************************************************************
'   Resource cache object pointers
'***************************************************************************

    Private mResourceCacheObject As clsResourceCacheObject

Public Sub Delete()

    intResource_Delete

End Sub

Public Property Set BITMAP(oNewValue As StdPicture)

    Set mBitmap = oNewValue

End Property

Public Property Get BITMAP() As StdPicture

    Set BITMAP = mBitmap

End Property

Public Property Let sBitmapFile(sNewValue As String)
    
    msBitmapFile = sNewValue

End Property

Public Property Get sBitmapFile() As String

    sBitmapFile = msBitmapFile

End Property


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

    Dim s As Integer
    
    Put #nFilenum, , CByte(GB_COLLISIONCODES)
    
    s = Len(msBitmapFile)
    Put #nFilenum, , s
    Put #nFilenum, , msBitmapFile

End Sub

Private Property Let intResource_ParentPath(RHS As String)

End Property

Private Property Get intResource_ParentPath() As String

End Property

Private Property Get intResource_ResourceType() As GB_FILETYPES

    intResource_ResourceType = GB_COLLISIONCODES

End Property

Private Function intResource_Unpack(nFilenum As Integer) As Boolean

    On Error GoTo HandleErrors

    Dim s As Integer
    
    intResource_Unpack = True
    
    Get #nFilenum, , s
    msBitmapFile = String(s, " ")
    Get #nFilenum, , msBitmapFile

    If msBitmapFile <> "" Then
        If Dir(msBitmapFile) <> GetTruncFilename(msBitmapFile) Then
            MsgBox "This file has a broken link to " & msBitmapFile & ".  Please find the missing file.", vbInformation, "Information"
            
            With mdiMain.Dialog
                .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
                .DialogTitle = "Locate " & GetTruncFilename(msBitmapFile)
                .Filename = ""
                .Filter = "All Files (*.*)|*.*"
                .ShowOpen
                If .Filename = "" Then
                    Exit Function
                End If
                gsCurPath = .Filename
                msBitmapFile = .Filename
            End With
        End If
        
        Set mBitmap = LoadPicture(msBitmapFile)
    End If

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBCollisionCodes:intResource_Unpack Error"
    intResource_Unpack = False
End Function


Private Sub intResource_UpdateClients(iClient As intResourceClient)

'***************************************************************************
'   Update all clients of this resource
'***************************************************************************
    
    If Not mResourceCacheObject Is Nothing Then
        mResourceCacheObject.UpdateClients iClient
    End If

End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    intResource_UpdateClients Nothing

End Sub


