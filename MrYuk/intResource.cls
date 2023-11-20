VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "intResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mResourceCacheObject As clsResourceCacheObject

Public Property Get CacheObject() As clsResourceCacheObject
End Property

Public Property Let ParentPath(sNewValue As String)
End Property

Public Property Get ParentPath() As String
End Property

Public Function Unpack(nFilenum As Integer) As Boolean
End Function

Public Sub Pack(nFilenum As Integer)
End Sub

Public Property Set CacheObject(oCacheObject As clsResourceCacheObject)
End Property

Public Sub UpdateClients(iClient As intResourceClient)
End Sub

Public Property Get ResourceType() As GB_FILETYPES
End Property

Public Sub Delete()
End Sub


