Attribute VB_Name = "modHotspot"
Option Explicit

Public Sub InitHotspot(MapName As MAP_NAME_CONSTANTS)

    Select Case MapName
    
        Case UNDGEYSER01
        
        Case UNDGEYSER02
        
        Case UNDGEYSER03
        
        Case UNDGEYSER04
        
        Case UNDGEYSER05
        
        Case UNDGEYSER06
        
        Case UNDGEYSER07
                
        Case UNDGEYSER08
        
        Case UNDGEYSER09
    
        Case UNDGEYSER10
    
            gMaps(MapName).AddHotspot 4, 8, DIR_DOWN
            gMaps(MapName).AddHotspot 5, 8, DIR_DOWN
            gMaps(MapName).AddHotspot 6, 8, DIR_DOWN
            gMaps(MapName).AddHotspot 7, 8, DIR_DOWN
    
        Case UNDGEYSER11
            
            gMaps(MapName).AddHotspot 4, 0, DIR_UP
            gMaps(MapName).AddHotspot 5, 0, DIR_UP
            gMaps(MapName).AddHotspot 6, 0, DIR_UP
            gMaps(MapName).AddHotspot 7, 0, DIR_UP
    
    End Select

End Sub

Public Sub RunHotspot(MapName As MAP_NAME_CONSTANTS, Hotspot As clsHotspot)

    Select Case MapName
    
        Case UNDGEYSER01
        
        Case UNDGEYSER02
        
        Case UNDGEYSER03
        
        Case UNDGEYSER04
        
        Case UNDGEYSER05
        
        Case UNDGEYSER06
        
        Case UNDGEYSER07
                
        Case UNDGEYSER08
        
        Case UNDGEYSER09
    
        Case UNDGEYSER10
    
            If Hotspot.xInTiles = 4 And Hotspot.yInTiles = 8 Then
                SelectMap UNDGEYSER11
                gPlayer.xInTiles = 4
                gPlayer.yInTiles = 0
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 5 And Hotspot.yInTiles = 8 Then
                SelectMap UNDGEYSER11
                gPlayer.xInTiles = 5
                gPlayer.yInTiles = 0
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 6 And Hotspot.yInTiles = 8 Then
                SelectMap UNDGEYSER11
                gPlayer.xInTiles = 6
                gPlayer.yInTiles = 0
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 7 And Hotspot.yInTiles = 8 Then
                SelectMap UNDGEYSER11
                gPlayer.xInTiles = 7
                gPlayer.yInTiles = 0
                UpdateScreen
            End If
    
        Case UNDGEYSER11
            
            If Hotspot.xInTiles = 4 And Hotspot.yInTiles = 0 Then
                SelectMap UNDGEYSER10
                gPlayer.xInTiles = 4
                gPlayer.yInTiles = 8
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 5 And Hotspot.yInTiles = 0 Then
                SelectMap UNDGEYSER10
                gPlayer.xInTiles = 5
                gPlayer.yInTiles = 8
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 6 And Hotspot.yInTiles = 0 Then
                SelectMap UNDGEYSER10
                gPlayer.xInTiles = 6
                gPlayer.yInTiles = 8
                UpdateScreen
            End If
    
            If Hotspot.xInTiles = 7 And Hotspot.yInTiles = 0 Then
                SelectMap UNDGEYSER10
                gPlayer.xInTiles = 7
                gPlayer.yInTiles = 8
                UpdateScreen
            End If
    
    End Select

End Sub


