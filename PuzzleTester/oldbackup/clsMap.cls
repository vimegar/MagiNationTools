VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moHotspots() As clsHotspot
Private mnHotspotCount As Integer

Private mIndex As MAP_NAME_CONSTANTS
Private mnTileWidth As Integer
Private mnTileHeight As Integer
Private miMapData() As Byte

Public Property Set Hotspots(Index As Integer, oNewValue As clsHotspot)

    Set moHotspots(Index) = oNewValue

End Property

Public Property Get Hotspots(Index As Integer) As clsHotspot

    Set Hotspots = moHotspots(Index)

End Property

Public Property Let Index(tNewValue As MAP_NAME_CONSTANTS)

    mIndex = tNewValue

End Property

Public Sub LoadMap(Filename As String)

    On Error GoTo HandleErrors

    If Filename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open Filename For Binary As #nFilenum
    
        Dim iFileType As Byte
        Dim nFormat As Integer
        Dim nPatStrLen As Integer
        Dim sPatStr As String
        
        Get #nFilenum, , iFileType
        Get #nFilenum, , nFormat
        Get #nFilenum, , nPatStrLen
        sPatStr = String(nPatStrLen, " ")
        Get #nFilenum, , sPatStr
        
        Get #nFilenum, , mnTileWidth
        Get #nFilenum, , mnTileHeight
        
        ResizeMap mnTileWidth, mnTileHeight
        
        Dim x As Integer
        Dim y As Integer
        Dim i As Byte
        
        For y = 0 To (mnTileHeight - 1)
            For x = 0 To (mnTileWidth - 1)
                Get #nFilenum, , i
                miMapData(x, y) = i
            Next x
        Next y
    
    Close #nFilenum

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modMaps:LoadMap Error"
End Sub


Public Property Get Index() As MAP_NAME_CONSTANTS

    Index = mIndex

End Property

Public Property Get MapData(x As Integer, y As Integer) As Byte

    MapData = miMapData(x, y)

End Property

Public Property Let MapData(x As Integer, y As Integer, iNewValue As Byte)

    miMapData(x, y) = iNewValue

End Property

Public Sub ResizeMap(Width As Integer, Height As Integer)

    ReDim miMapData(Width, Height)

End Sub


Public Property Get TileHeight() As Integer
    
    TileHeight = mnTileHeight

End Property

Public Property Let TileHeight(nNewValue As Integer)

    mnTileHeight = nNewValue

End Property

Public Property Get TileWidth() As Integer
    
    TileWidth = mnTileWidth

End Property

Public Property Let TileWidth(nNewValue As Integer)

    mnTileWidth = nNewValue

End Property



Public Sub AddHotspot(xTile As Integer, yTile As Integer, Code As Integer)

    mnHotspotCount = mnHotspotCount + 1

    ReDim Preserve moHotspots(mnHotspotCount - 1)
    Set moHotspots(mnHotspotCount - 1) = New clsHotspot
    
    moHotspots(mnHotspotCount - 1).xTile = xTile
    moHotspots(mnHotspotCount - 1).yTile = yTile
    moHotspots(mnHotspotCount - 1).nCode = Code

End Sub

Public Property Get HotspotCount() As Integer

    HotspotCount = mnHotspotCount
    
End Property
Public Sub RunHotspot(Hotspot As clsHotspot)

    Select Case mIndex

        Case UNDGEYSER10
            
            Select Case Hotspot.nCode
            
            
            End Select

        Case UNDGEYSER11
            
            Select Case Hotspot.nCode
            
                Case 10
                    
                    If (Hotspot.xTile = 0 And Hotspot.yTile = 7) Or _
                       (Hotspot.xTile = 0 And Hotspot.yTile = 8) Or _
                       (Hotspot.xTile = 0 And Hotspot.yTile = 9) Or _
                       (Hotspot.xTile = 0 And Hotspot.yTile = 10) Then
                    
                        gTony.xPixel = gTony.xPixel + 128
                    
                    End If
                    
                    If (Hotspot.xTile = 4 And Hotspot.yTile = 0) Or _
                       (Hotspot.xTile = 5 And Hotspot.yTile = 0) Or _
                       (Hotspot.xTile = 6 And Hotspot.yTile = 0) Or _
                       (Hotspot.xTile = 7 And Hotspot.yTile = 0) Then
                    
                        gnCurrentMap = UNDGEYSER10
                    
                    End If
                    
                    If (Hotspot.xTile = 27 And Hotspot.yTile = 7) Or _
                       (Hotspot.xTile = 27 And Hotspot.yTile = 8) Or _
                       (Hotspot.xTile = 27 And Hotspot.yTile = 9) Or _
                       (Hotspot.xTile = 27 And Hotspot.yTile = 10) Then
                    
                        gTony.xPixel = gTony.xPixel - 128
                    
                    End If
            
            End Select
        
    End Select

End Sub

Public Sub InitHotspots()

    Select Case mIndex
        
        Case UNDGEYSER10
        
        Case UNDGEYSER11
            
            AddHotspot 0, 7, 10
            AddHotspot 0, 8, 10
            AddHotspot 0, 9, 10
            AddHotspot 0, 10, 10
            
            AddHotspot 4, 0, 10
            AddHotspot 5, 0, 10
            AddHotspot 6, 0, 10
            AddHotspot 7, 0, 10
        
            AddHotspot 27, 7, 10
            AddHotspot 27, 8, 10
            AddHotspot 27, 9, 10
            AddHotspot 27, 10, 10
            
    End Select
    
End Sub



