Attribute VB_Name = "modSwitches"
Option Explicit

Public Type tSwitch
    xTile As Integer
    yTile As Integer
    nCode As Integer
End Type

Public gSwitches() As tSwitch

Private mnSwitchCount As Integer
Public Sub AddSwitch(xTile As Integer, yTile As Integer, Code As Integer)

    mnSwitchCount = mnSwitchCount + 1

    ReDim Preserve gSwitches(mnSwitchCount - 1)
    
    gSwitches(mnSwitchCount - 1).xTile = xTile
    gSwitches(mnSwitchCount - 1).yTile = yTile
    gSwitches(mnSwitchCount - 1).nCode = Code

End Sub


Public Property Get SwitchCount() As Integer

    SwitchCount = mnSwitchCount

End Property

Public Sub FlipSwitch(Switch As tSwitch)

    Select Case Switch.nCode
    
        Case 6, 7, 8, 9
            
            gnBlockDownID = Switch.nCode - 4
            
        Case 10 'map link
    
            MsgBox "map link"
    
    End Select

End Sub
