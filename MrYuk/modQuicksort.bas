Attribute VB_Name = "modQuicksort"
Option Explicit

Public Enum QUICKSORT_COMPARE_TYPES
    GreaterThanOrEqualTo = 0
    LessThanOrEqualTo = 1
End Enum

Public Sub QuicksortCTSubtile(oArray() As clsCTSubtile, nStart As Integer, nEnd As Integer)

    On Error GoTo HandleErrors

    Dim q As Integer

    If nStart < nEnd Then
        q = mPartitionCTSubtile(oArray, nStart, nEnd)
        QuicksortCTSubtile oArray, nStart, q
        QuicksortCTSubtile oArray, q + 1, nEnd
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modQuicksort:QuicksortCTSubtile Error"
End Sub

Private Function mPartitionCTSubtile(oArray() As clsCTSubtile, nStart As Integer, nEnd As Integer) As Integer

    On Error GoTo HandleErrors

    Dim X As clsCTSubtile
    Dim d As clsCTSubtile
    Dim i As Integer
    Dim j As Integer
    
    Set X = oArray(nStart)
    i = nStart - 1
    j = nEnd + 1
    
    While True
        
        Do
            If j = nStart Then
                Exit Do
            Else
                j = j - 1
            End If
        Loop Until oArray(j).Compare(X, LessThanOrEqualTo)
        
        Do
            If i = nEnd Then
                Exit Do
            Else
                i = i + 1
            End If
        Loop Until oArray(i).Compare(X, GreaterThanOrEqualTo)
        
        If i < j Then
            Set d = oArray(j)
            Set oArray(j) = oArray(i)
            Set oArray(i) = d
        Else
            mPartitionCTSubtile = j
            Exit Function
        End If
        
    Wend

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "modQuicksort:mPartition Error"
End Function

Public Sub QuicksortBPFile(oArray() As clsBPFile, nStart As Integer, nEnd As Integer)

    Dim q As Integer

    If nStart < nEnd Then
        
        q = mPartitionBPFile(oArray, nStart, nEnd)
        
        QuicksortBPFile oArray, nStart, q
        QuicksortBPFile oArray, q + 1, nEnd
        
    End If

End Sub

Public Sub QuicksortHSRect(oArray() As clsHotspotRect, nStart As Integer, nEnd As Integer)

    Dim q As Integer

    If nStart < nEnd Then
        
        q = mPartitionHSRect(oArray, nStart, nEnd)
        
        QuicksortHSRect oArray, nStart, q
        QuicksortHSRect oArray, q + 1, nEnd
        
    End If

End Sub

Private Function mPartitionBPFile(oArray() As clsBPFile, nStart As Integer, nEnd As Integer) As Integer

    Dim dummy As clsBPFile
    Dim X As clsBPFile
    Dim i As Integer
    Dim j As Integer

    Set X = oArray(nStart)
    i = nStart - 1
    j = nEnd + 1

    While True
        
        Do
            If j = nStart Then
                Exit Do
            Else
                j = j - 1
            End If
        Loop Until oArray(j).FileSize <= X.FileSize

        Do
            If i = nEnd Then
                Exit Do
            Else
                i = i + 1
            End If
        Loop Until oArray(i).FileSize >= X.FileSize
        
        If i < j Then
            Set dummy = oArray(i)
            Set oArray(i) = oArray(j)
            Set oArray(j) = dummy
        Else
            mPartitionBPFile = j
            Exit Function
        End If

    Wend
    
End Function

Private Function mPartitionHSRect(oArray() As clsHotspotRect, nStart As Integer, nEnd As Integer) As Integer

    Dim dummy As clsHotspotRect
    Dim X As clsHotspotRect
    Dim i As Integer
    Dim j As Integer

    Set X = oArray(nStart)
    i = nStart - 1
    j = nEnd + 1

    While True
        
        Do
            If j = nStart Then
                Exit Do
            Else
                j = j - 1
            End If
        Loop Until oArray(j).nType <= X.nType

        Do
            If i = nEnd Then
                Exit Do
            Else
                i = i + 1
            End If
        Loop Until oArray(i).nType >= X.nType
        
        If i < j Then
            Set dummy = oArray(i)
            Set oArray(i) = oArray(j)
            Set oArray(j) = dummy
        Else
            mPartitionHSRect = j
            Exit Function
        End If

    Wend
    
End Function

