VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCTSubtile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements intQuicksortObj

Private mlPixelData(4) As Long
Private mbXFlip As Boolean
Private mbYFlip As Boolean
Private moParent As clsCT

Public Function Compare(CompareToObj As clsCTSubtile, CompareType As QUICKSORT_COMPARE_TYPES) As Boolean
    
    Compare = intQuicksortObj_Compare(CompareToObj, CompareType)

End Function



Public Property Set Parent(oNewValue As clsCT)

    Set moParent = oNewValue

End Property

Public Property Get Parent() As clsCT

    Set Parent = moParent

End Property

Public Property Let PixelData(Index As Integer, lNewValue As Long)

    mlPixelData(Index) = lNewValue

End Property

Public Property Get PixelData(Index As Integer) As Long

    PixelData = mlPixelData(Index)

End Property

Public Property Get XFlip() As Boolean

    XFlip = mbXFlip

End Property

Public Property Let XFlip(bNewValue As Boolean)

    mbXFlip = bNewValue

End Property

Public Property Get YFlip() As Boolean

    YFlip = mbYFlip

End Property

Public Property Let YFlip(bNewValue As Boolean)

    mbYFlip = bNewValue

End Property

Private Function intQuicksortObj_Compare(CompareToObj As clsCTSubtile, CompareType As QUICKSORT_COMPARE_TYPES) As Boolean

    Select Case CompareType
    
        Case LessThanOrEqualTo
        
            If PixelData(0) < CompareToObj.PixelData(0) Then
                intQuicksortObj_Compare = True
                Exit Function
            ElseIf PixelData(0) = CompareToObj.PixelData(0) Then
                
                If PixelData(1) < CompareToObj.PixelData(1) Then
                    intQuicksortObj_Compare = True
                    Exit Function
                ElseIf PixelData(1) = CompareToObj.PixelData(1) Then
                    
                    If PixelData(2) < CompareToObj.PixelData(2) Then
                        intQuicksortObj_Compare = True
                        Exit Function
                    ElseIf PixelData(2) = CompareToObj.PixelData(2) Then
                        
                        If PixelData(3) < CompareToObj.PixelData(3) Then
                            intQuicksortObj_Compare = True
                            Exit Function
                        ElseIf PixelData(3) = CompareToObj.PixelData(3) Then
                            intQuicksortObj_Compare = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            intQuicksortObj_Compare = False
        
        Case GreaterThanOrEqualTo
        
            If PixelData(0) > CompareToObj.PixelData(0) Then
                intQuicksortObj_Compare = True
                Exit Function
            ElseIf PixelData(0) = CompareToObj.PixelData(0) Then
                
                If PixelData(1) > CompareToObj.PixelData(1) Then
                    intQuicksortObj_Compare = True
                    Exit Function
                ElseIf PixelData(1) = CompareToObj.PixelData(1) Then
                    
                    If PixelData(2) > CompareToObj.PixelData(2) Then
                        intQuicksortObj_Compare = True
                        Exit Function
                    ElseIf PixelData(2) = CompareToObj.PixelData(2) Then
                        
                        If PixelData(3) > CompareToObj.PixelData(3) Then
                            intQuicksortObj_Compare = True
                            Exit Function
                        ElseIf PixelData(3) = CompareToObj.PixelData(3) Then
                            intQuicksortObj_Compare = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            intQuicksortObj_Compare = False

    End Select

End Function


