Attribute VB_Name = "modGeyser2"
Option Explicit

Private Const NUM_SWITCHES = 5

Public Enum SWITCH_TYPE_CONSTANTS
    SWITCH_BLUE = 0
    SWITCH_GREEN = 1
    SWITCH_YELLOW = 2
    SWITCH_RED = 3
End Enum

Private Type SWITCH_LOC
    xInTiles As Integer
    yInTiles As Integer
    tType As SWITCH_TYPE_CONSTANTS
End Type

Private mtSwitchLocs(NUM_SWITCHES - 1) As SWITCH_LOC
Public Sub FlipColorSwitch(MapName As MAP_NAME_CONSTANTS, xInTiles As Integer, yInTiles As Integer)

    Select Case MapName
    
        Case UNDGEYSER11
            
            If xInTiles = 9 And yInTiles = 2 Then
                MsgBox "flipped blue switch"
            End If
            
    End Select

End Sub


Public Sub InitGeyser2()

'Define blue switches
    

End Sub


