Attribute VB_Name = "modSwitch"
Option Explicit

Public Sub InitSwitch(MapName As MAP_NAME_CONSTANTS)

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
    
        Case UNDGEYSER11
            
            gMaps(MapName).AddSwitch 9, 2
    
    End Select

End Sub

Public Sub RunSwitch(MapName As MAP_NAME_CONSTANTS, Switch As clsSwitch)

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
    
        Case UNDGEYSER11
            
            If Switch.xInTiles = 9 And Switch.yInTiles = 2 Then
            
                FlipColorSwitch MapName, Switch.xInTiles, Switch.yInTiles
                         
            End If
    
    End Select

End Sub


