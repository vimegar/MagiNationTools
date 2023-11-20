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

Private mnWidthInTiles As Integer
Private mnHeightInTiles As Integer
Private miMapData() As Byte

Private moHotspots() As clsHotspot
Private mnHotspotCount As Integer

Private moSwitches() As clsSwitch
Private mnSwitchCount As Integer

Private mtMapName As MAP_NAME_CONSTANTS

Public Sub AddHotspot(xInTiles As Integer, yInTiles As Integer, Direction As DIRECTION_CONSTANTS)

    mnHotspotCount = mnHotspotCount + 1
    ReDim Preserve moHotspots(mnHotspotCount - 1)
    Set moHotspots(mnHotspotCount - 1) = New clsHotspot
    
    moHotspots(mnHotspotCount - 1).xInTiles = xInTiles
    moHotspots(mnHotspotCount - 1).yInTiles = yInTiles
    moHotspots(mnHotspotCount - 1).Direction = Direction
    
End Sub


Public Sub AddSwitch(xInTiles As Integer, yInTiles As Integer)

    mnSwitchCount = mnSwitchCount + 1
    ReDim Preserve moSwitches(mnSwitchCount - 1)
    Set moSwitches(mnSwitchCount - 1) = New clsSwitch
    
    moSwitches(mnSwitchCount - 1).xInTiles = xInTiles
    moSwitches(mnSwitchCount - 1).yInTiles = yInTiles

End Sub

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
        
        Get #nFilenum, , mnWidthInTiles
        Get #nFilenum, , mnHeightInTiles
        
        ResizeMap mnWidthInTiles, mnHeightInTiles
        
        Dim x As Integer
        Dim y As Integer
        Dim i As Byte
        
        For y = 0 To (mnHeightInTiles - 1)
            For x = 0 To (mnWidthInTiles - 1)
                Get #nFilenum, , i
                miMapData(x, y) = i
            Next x
        Next y
    
    Close #nFilenum

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsMap:LoadMap Error"
End Sub

Public Property Let MapName(tNewValue As MAP_NAME_CONSTANTS)

    mtMapName = tNewValue

End Property

Public Property Get MapName() As MAP_NAME_CONSTANTS

    MapName = mtMapName

End Property

Public Sub ResizeMap(Width As Integer, Height As Integer)

    ReDim miMapData(Width - 1, Height - 1)

End Sub

Public Property Get HotspotCount() As Integer

    HotspotCount = mnHotspotCount

End Property

Public Property Get Hotspots(Index As Integer) As clsHotspot

    Set Hotspots = moHotspots(Index)

End Property

Public Property Set Hotspots(Index As Integer, oNewValue As clsHotspot)

    Set moHotspots(Index) = oNewValue

End Property

Public Property Get MapData(x As Integer, y As Integer) As Byte
    
    MapData = miMapData(x, y)
    
End Property

Public Property Let MapData(x As Integer, y As Integer, iNewValue As Byte)

    miMapData(x, y) = iNewValue

End Property


Public Property Get SwitchCount() As Integer

    SwitchCount = mnSwitchCount

End Property

Public Property Set Switches(Index As Integer, oNewValue As clsSwitch)

    Set moSwitches(Index) = oNewValue

End Property

Public Property Get Switches(Index As Integer) As clsSwitch

    Set Switches = moSwitches(Index)

End Property

Public Property Get WidthInTiles() As Integer

    WidthInTiles = mnWidthInTiles

End Property

Public Property Let WidthInTiles(nNewValue As Integer)

    mnWidthInTiles = nNewValue

End Property

Public Property Get HeightInTiles() As Integer

    HeightInTiles = mnHeightInTiles

End Property

Public Property Let HeightInTiles(nNewValue As Integer)

    mnHeightInTiles = nNewValue

End Property
