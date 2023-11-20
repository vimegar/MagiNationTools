VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResourceCacheObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFilename As String
Private mClientList As New Collection
Private mData As intResource

Public Property Get Data() As intResource

    Set Data = mData

End Property

Public Property Set Data(oNewValue As intResource)

    Set mData = oNewValue

End Property

Public Sub Delete()

    If Not mData Is Nothing Then
        mData.Delete
    End If

End Sub

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get ClientList() As Collection

    Set ClientList = mClientList

End Property

Public Sub AddClient(iClient As intResourceClient)
    
'***************************************************************************
'   Add a client to the resource
'***************************************************************************
    
    Dim clientCompare As intResourceClient
    
    For Each clientCompare In ClientList
        If clientCompare Is iClient Then
            Exit Sub
        End If
    Next clientCompare
    
    ClientList.Add iClient
    
End Sub

Public Sub RemoveClient(iClient As intResourceClient)
        
'***************************************************************************
'   Remove client from the resources client list
'***************************************************************************
        
    Dim i As Integer
    
    For i = 1 To ClientList.count
        If i > ClientList.count Then Exit For
        If ClientList(i) Is iClient Then
            ClientList.remove i
            If i <= ClientList.count Then
                i = i - 1
            End If
        End If
    Next i
    
End Sub

Public Sub UpdateClients(iClient As intResourceClient)

'***************************************************************************
'   Call the Update() event for each client of this resource
'***************************************************************************

    Dim i As Integer
    
    For i = 1 To ClientList.count
        If Not ClientList(i) Is iClient Then
            ClientList(i).Update GB_RESOURCECHANGED
        End If
    Next i

End Sub

