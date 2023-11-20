Attribute VB_Name = "modColCodes"
Option Explicit

Public gColCodes(NUM_TILES - 1) As clsColCodes
Public Sub InitColCodes()

    Set gColCodes(0) = New clsColCodes
    gColCodes(0).Walkable = False
    
    Set gColCodes(1) = New clsColCodes
    gColCodes(1).Walkable = True
    
    Set gColCodes(2) = New clsColCodes
    gColCodes(2).Walkable = False
    
    Set gColCodes(3) = New clsColCodes
    gColCodes(3).Walkable = False
    
    Set gColCodes(4) = New clsColCodes
    gColCodes(4).Walkable = False
    
    Set gColCodes(5) = New clsColCodes
    gColCodes(5).Walkable = False
    
    Set gColCodes(6) = New clsColCodes
    gColCodes(6).Walkable = False
    
    Set gColCodes(7) = New clsColCodes
    gColCodes(7).Walkable = False
    
    Set gColCodes(8) = New clsColCodes
    gColCodes(8).Walkable = False
    
    Set gColCodes(9) = New clsColCodes
    gColCodes(9).Walkable = False
    
    Set gColCodes(10) = New clsColCodes
    gColCodes(10).Walkable = True
    
    Set gColCodes(11) = New clsColCodes
    gColCodes(11).Walkable = True
    
    Set gColCodes(12) = New clsColCodes
    gColCodes(12).Walkable = True
    
    Set gColCodes(13) = New clsColCodes
    gColCodes(13).Walkable = True
    
    Set gColCodes(14) = New clsColCodes
    gColCodes(14).Walkable = True
    
End Sub


