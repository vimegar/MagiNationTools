Attribute VB_Name = "modRLE"
 Option Explicit

Public Enum SPAN_TYPES
    SPAN_NORM = 0
    SPAN_NORMGRANDE = 1
    SPAN_RLE = 2
    SPAN_RLEGRANDE = 3
End Enum

Public Type tSpan
    nType As SPAN_TYPES
    Length As Integer
    iBytes() As Byte
End Type

Public mnDebug As Integer

Public Sub AddSpanRLE(nSpanLen As Integer, iFile() As Byte, CurrentPos As Integer, OutputFilenum As Integer)
    
    If nSpanLen < 64 Then
        'NORMAL
        Put #OutputFilenum, , CByte(BitStringToNum("10" & NumToBitString(nSpanLen, 6)))
        Put #OutputFilenum, , CByte(iFile(CurrentPos))
    Else
        'GRANDE
        Put #OutputFilenum, , CByte(BitStringToNum("11" & Mid$(NumToBitString(nSpanLen, 14), 1, 6)))
        Put #OutputFilenum, , CByte(BitStringToNum(Mid$(NumToBitString(nSpanLen, 14), 7)))
        Put #OutputFilenum, , CByte(iFile(CurrentPos))
    End If

    CurrentPos = CurrentPos + nSpanLen

End Sub

Public Sub AddSpanNormal(nSpanLen As Integer, iFile() As Byte, CurrentPos As Integer, OutputFilenum As Integer)

    On Error GoTo HandleErrors

    If nSpanLen < 64 Then
        'NORMAL
        Put #OutputFilenum, , CByte(BitStringToNum("00" & NumToBitString(nSpanLen, 6)))
    Else
        'GRANDE
        Put #OutputFilenum, , CByte(BitStringToNum("01" & Mid$(NumToBitString(nSpanLen, 14), 1, 6)))
        Put #OutputFilenum, , CByte(BitStringToNum(Mid$(NumToBitString(nSpanLen, 14), 7)))
    End If

    Dim i As Integer

    For i = 0 To (nSpanLen - 1)
        Put #OutputFilenum, , CByte(iFile(CurrentPos + i))
    Next i

    CurrentPos = CurrentPos + nSpanLen

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:AddSpanNormal Error"
End Sub

Public Sub DebugPrint(Filename As String, SpanArray() As tSpan)

    Dim i As Integer
    Dim str As String
    Dim nFilenum As Integer
    nFilenum = FreeFile
    Open Filename For Output As #nFilenum
        For i = 0 To (UBound(SpanArray) - 1)
            Select Case SpanArray(i).nType
                Case SPAN_NORM
                    str = "[Norm"
                Case SPAN_RLE
                    str = "[RLE"
                Case SPAN_NORMGRANDE
                    str = "[Norm Grande"
                Case SPAN_RLEGRANDE
                    str = "[RLE Grande"
            End Select
        
            Print #nFilenum, str & vbTab & Format$(CStr(SpanArray(i).Length), "000") & " ]" & CStr(i)
        
        Next i
    Close #nFilenum

End Sub

Public Function GetInputFile(InputFilenum As Integer, iFile() As Byte, CompressLen As Integer, HeaderSize As Integer) As Integer

    Dim nCount As Integer
    nCount = 0
    
    Dim nFileLen As Integer
    Dim nFileLoc As Integer
    
    nFileLen = LOF(InputFilenum)

    If CompressLen = 0 Then
        CompressLen = nFileLen - loc(InputFilenum)
    End If

    ReDim iFile(nFileLen - 1) As Byte

    Seek #InputFilenum, (HeaderSize + 1)

    Do Until (nCount = nFileLen)
        Get #InputFilenum, , iFile(nCount)
        nCount = nCount + 1
    Loop
    
    GetInputFile = CompressLen
    
End Function

Public Function GetSpanLen(bRLE As Boolean, iFile() As Byte, ByVal CurrentPos As Integer) As Integer

    On Error GoTo HandleErrors

    If bRLE Then
        GetSpanLen = 1
    Else
        GetSpanLen = 0
    End If

    Do Until (bRLE = (iFile(CurrentPos) <> iFile(CurrentPos + 1)))
        
        CurrentPos = CurrentPos + 1
        GetSpanLen = GetSpanLen + 1
        
        If (CurrentPos + 1) > UBound(iFile) Then
            Exit Do
        End If
        
    Loop
   
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:GetSpanLen Error"
End Function
Public Sub CreateSpanTable(OutputFilenum As Integer, HeaderSize As Integer, SpanArray() As tSpan)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim iByte As Byte
    Dim iByte2 As Byte
    Dim SpanArrayIndex As Integer
    Dim dummy As Byte
    Dim bExitFlag As Boolean

    Seek #OutputFilenum, 1
    
    'get past header info
    For i = 0 To (HeaderSize - 1)
        Get #OutputFilenum, , dummy
    Next i
    
    Do Until EOF(OutputFilenum)
        
        'read in byte
        Get #OutputFilenum, , iByte
                
        If iByte = 0 Then
            bExitFlag = True
        End If
            
        'add new array element
        ReDim Preserve SpanArray(UBound(SpanArray) + 1)
        SpanArrayIndex = UBound(SpanArray) - 1
        
        'get type
        SpanArray(SpanArrayIndex).nType = BitStringToNum("000000" & Mid$(NumToBitString(iByte), 1, 2))
        
        'get len
        Select Case SpanArray(SpanArrayIndex).nType
            Case SPAN_NORM
                
                SpanArray(SpanArrayIndex).Length = BitStringToNum(Mid$(NumToBitString(iByte), 3, 6))
                
                ReDim SpanArray(SpanArrayIndex).iBytes(SpanArray(SpanArrayIndex).Length)
                For i = 0 To (SpanArray(SpanArrayIndex).Length - 1)
                    Get #OutputFilenum, , dummy
                    SpanArray(SpanArrayIndex).iBytes(i) = dummy
                Next i
                
            Case SPAN_RLE
                
                SpanArray(SpanArrayIndex).Length = BitStringToNum(Mid$(NumToBitString(iByte), 3, 6))
                
                ReDim SpanArray(SpanArrayIndex).iBytes(1)
                Get #OutputFilenum, , dummy
                SpanArray(SpanArrayIndex).iBytes(0) = dummy
                
            Case SPAN_NORMGRANDE
                
                Get #OutputFilenum, , iByte2
                
                SpanArray(SpanArrayIndex).Length = BitStringToNum(Mid$(NumToBitString(iByte), 3, 6) & NumToBitString(iByte2))
                
                ReDim SpanArray(SpanArrayIndex).iBytes(SpanArray(SpanArrayIndex).Length)
                For i = 0 To (SpanArray(SpanArrayIndex).Length - 1)
                    Get #OutputFilenum, , dummy
                    SpanArray(SpanArrayIndex).iBytes(i) = dummy
                Next i
                
            Case SPAN_RLEGRANDE
                
                Get #OutputFilenum, , iByte2
                
                SpanArray(SpanArrayIndex).Length = BitStringToNum(Mid$(NumToBitString(iByte), 3, 6) & NumToBitString(iByte2))
                
                ReDim SpanArray(SpanArrayIndex).iBytes(1)
                Get #OutputFilenum, , dummy
                SpanArray(SpanArrayIndex).iBytes(0) = dummy
                
        End Select
        
        If bExitFlag Then
            Exit Do
        End If
        
    Loop
    
    'ReDim Preserve SpanArray(UBound(SpanArray) - 1)

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:CreateSpanTable Error"
End Sub

Public Sub OutputNormGrandeFile(InputFilenum As Integer, OutputFilename As String, OutputFilenum As Integer, HeaderSize As Integer)

    Dim i As Integer
    Dim str As String
    Dim d As Byte
    
    Close #OutputFilenum
    
    DeleteFile OutputFilename
    
    Dim nFilenum As String
    OutputFilenum = FreeFile
    
    Open OutputFilename For Binary As #OutputFilenum
        
    Seek #InputFilenum, 1

    For i = 0 To (HeaderSize - 1)
        Get #InputFilenum, , d
        Put #OutputFilenum, , d
    Next i

    str = NumToBitString(LOF(InputFilenum) - HeaderSize, 14)
    Put #OutputFilenum, , CByte(BitStringToNum("01" & Mid$(str, 1, 6)))
    Put #OutputFilenum, , CByte(BitStringToNum(Mid$(str, 7, 14)))
    
    Do Until EOF(InputFilenum)
        Get #InputFilenum, , d
        Put #OutputFilenum, , d
    Loop

End Sub

Public Sub OutputOptimizeFile(InputFilenum As Integer, OutputFilenum As Integer, HeaderSize As Integer, SpanArray() As tSpan, OutputFilename As String)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim iByte As Byte
    Dim iByte2 As Byte
    Dim lenStr As String
    Dim SpanCursor As Integer

    Close #OutputFilenum
    DeleteFile OutputFilename
    Open OutputFilename For Binary As #OutputFilenum

    'start at beginning of original file
    Seek #InputFilenum, 1
    
    For i = 0 To (HeaderSize - 1)
        Get #InputFilenum, , iByte
        Put #OutputFilenum, , iByte
    Next i
    
    For SpanCursor = 0 To (UBound(SpanArray) - 1)
        
        Select Case SpanArray(SpanCursor).nType
            Case SPAN_NORM
                
                iByte = BitStringToNum(Mid$(NumToBitString(SpanArray(SpanCursor).nType), 7, 2) & Mid$(NumToBitString(SpanArray(SpanCursor).Length), 3, 6))
                Put #OutputFilenum, , iByte
                
                For i = 0 To UBound(SpanArray(SpanCursor).iBytes) - 1
                    iByte = SpanArray(SpanCursor).iBytes(i)
                    Put #OutputFilenum, , iByte
                Next i
                
            Case SPAN_RLE
            
                iByte = BitStringToNum(Mid$(NumToBitString(SpanArray(SpanCursor).nType), 7, 2) & Mid$(NumToBitString(SpanArray(SpanCursor).Length), 3, 6))
                Put #OutputFilenum, , iByte
                
                Put #OutputFilenum, , SpanArray(SpanCursor).iBytes(0)
            
            Case SPAN_NORMGRANDE
            
                lenStr = NumToBitString(SpanArray(SpanCursor).Length, 14)
                
                iByte = BitStringToNum(Mid$(NumToBitString(SpanArray(SpanCursor).nType), 7, 2) & Mid$(lenStr, 1, 6))
                iByte2 = BitStringToNum(Mid$(lenStr, 7, 8))
                
                Put #OutputFilenum, , iByte
                Put #OutputFilenum, , iByte2
                
                For i = 0 To UBound(SpanArray(SpanCursor).iBytes) - 1
                    iByte = SpanArray(SpanCursor).iBytes(i)
                    Put #OutputFilenum, , iByte
                Next i
            
            Case SPAN_RLEGRANDE
            
                lenStr = NumToBitString(SpanArray(SpanCursor).Length, 14)
                
                iByte = BitStringToNum(Mid$(NumToBitString(SpanArray(SpanCursor).nType), 7, 2) & Mid$(lenStr, 1, 6))
                iByte2 = BitStringToNum(Mid$(lenStr, 7, 8))
                
                Put #OutputFilenum, , iByte
                Put #OutputFilenum, , iByte2
                
                Put #OutputFilenum, , SpanArray(SpanCursor).iBytes(0)
            
        End Select
        
    Next SpanCursor
    
    Put #OutputFilenum, , CByte(0)

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:OutputOptimizeFile Error"
End Sub

Public Sub RemoveDoubleNorms(SpanArray() As tSpan)

    On Error GoTo HandleErrors

    Dim SpanCursor As Integer
    
    For SpanCursor = 0 To (UBound(SpanArray) - 1)
    
        If SpanCursor > (UBound(SpanArray) - 1) Then
            Exit For
        End If
    
        If ((SpanArray(SpanCursor).nType = SPAN_NORM) Or (SpanArray(SpanCursor).nType = SPAN_NORMGRANDE)) _
            And ((SpanArray(SpanCursor + 1).nType = SPAN_NORM) Or (SpanArray(SpanCursor + 1).nType = SPAN_NORMGRANDE)) Then
        
            Dim j As Integer
            For j = (SpanArray(SpanCursor).Length + 1) To (SpanArray(SpanCursor + 1).Length + (SpanArray(SpanCursor).Length + 1)) - 1
                ReDim Preserve SpanArray(SpanCursor).iBytes(j)
                SpanArray(SpanCursor).iBytes(j - 1) = SpanArray(SpanCursor + 1).iBytes(j - (SpanArray(SpanCursor).Length + 1))
            Next j

            SpanArray(SpanCursor).Length = SpanArray(SpanCursor).Length + SpanArray(SpanCursor + 1).Length
        
            If SpanArray(SpanCursor).Length < 64 Then
                SpanArray(SpanCursor).nType = SPAN_NORM
            Else
                SpanArray(SpanCursor).nType = SPAN_NORMGRANDE
            End If

            For j = (SpanCursor + 2) To (UBound(SpanArray) - 1)
                SpanArray(j - 1) = SpanArray(j)
            Next j
            ReDim Preserve SpanArray(UBound(SpanArray) - 1)

        End If
        
        
    
    Next SpanCursor
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:RemoveDoubleNorms Error"
End Sub

Public Sub RLENormGrande(InputFilenum As Integer, OutputFilenum As Integer, HeaderSize As Integer, Optional CompressLen As Integer)

    On Error GoTo HandleErrors

    Dim dummy As Integer
    Dim iFile() As Byte

    CompressLen = GetInputFile(InputFilenum, iFile, CompressLen, HeaderSize)

    AddSpanNormal UBound(iFile) + 1, iFile, dummy, OutputFilenum
    
    Put #OutputFilenum, , CByte(0)

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modMain:RLENormGrande Error"
End Sub

Public Sub RLEStandard(InputFilenum As Integer, OutputFilenum As Integer, HeaderSize As Integer, Optional CompressLen As Integer)

    On Error GoTo HandleErrors

    Dim iFile() As Byte
        
    CompressLen = GetInputFile(InputFilenum, iFile, CompressLen, HeaderSize)

    Dim bRLE As Boolean
    Dim nSpanLen As Integer
    Dim CurrentPos As Integer
    
    CurrentPos = 0
    
    Do Until CurrentPos >= CompressLen
    
        If (CurrentPos + 1) > UBound(iFile) Then
            nSpanLen = 1
            AddSpanNormal nSpanLen, iFile, CurrentPos, OutputFilenum
            Exit Do
        End If
        
        bRLE = (iFile(CurrentPos) = iFile((CurrentPos + 1)))
        
        nSpanLen = GetSpanLen(bRLE, iFile, CurrentPos)
        
        If bRLE Then
            AddSpanRLE nSpanLen, iFile, CurrentPos, OutputFilenum
        Else
            AddSpanNormal nSpanLen, iFile, CurrentPos, OutputFilenum
        End If
        
    Loop
          
    Put #OutputFilenum, , CByte(0)
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modMain:RLEStandard Error"
End Sub
    
Public Sub PackRLE(InputFilename As String, OutputFilename As String, HeaderSize As Integer, Optional Progress As ProgressBar)

    On Error GoTo HandleErrors
    
'input
    Dim nInputFilenum As Integer
    nInputFilenum = FreeFile
    
    Open InputFilename For Binary As #nInputFilenum
    
    Dim i As Integer
    ReDim iBytes(HeaderSize) As Byte

    For i = 0 To (HeaderSize - 1)
        Get #nInputFilenum, , iBytes(i)
    Next i
    
'output
    Dim str As String
    Dim tempNum1 As Integer
    Dim tempNum2 As Integer
    Dim tempNum3 As Integer
    
    For i = Len(OutputFilename) To 1 Step -1
        If Mid$(OutputFilename, i, 1) = "\" Then
            str = Mid$(OutputFilename, 1, i)
            Exit For
        End If
    Next i
    
    DeleteFile str & "temp1.bin"
    DeleteFile str & "temp2.bin"
    DeleteFile str & "temp3.bin"
    
    tempNum1 = FreeFile
    Open str & "temp1.bin" For Binary As #tempNum1
    
    tempNum2 = FreeFile
    Open str & "temp2.bin" For Binary As #tempNum2
    
    For i = 0 To (HeaderSize - 1)
        Put #tempNum1, , iBytes(i)
        Put #tempNum2, , iBytes(i)
    Next i
    
    RLEStandard nInputFilenum, tempNum1, HeaderSize
    Close #tempNum1
    
    RLENormGrande nInputFilenum, tempNum2, HeaderSize
    Close #tempNum2
    
    CopyFile str & "temp1.bin", str & "temp3.bin", False
    tempNum3 = FreeFile
    Open str & "temp3.bin" For Binary As #tempNum3
    RLENoDither nInputFilenum, tempNum3, HeaderSize, str & "temp3.bin"
    Close #tempNum3
    
    Close #nInputFilenum
        
    'check for smallest file
    Dim a As Long
    Dim b As Long
    Dim c As Long
    
    a = FileLen(str & "temp1.bin")
    b = FileLen(str & "temp2.bin")
    c = FileLen(str & "temp3.bin")
    
    If (a <= b) And (a <= c) Then
        CopyFile str & "temp1.bin", str & GetTruncFilename(OutputFilename), False
    ElseIf (b <= a) And (b <= c) Then
        CopyFile str & "temp2.bin", str & GetTruncFilename(OutputFilename), False
    Else
        'MsgBox "The file " & InputFilename & " has been output using the RLENoDither method.", vbInformation, "Information"
        CopyFile str & "temp3.bin", str & GetTruncFilename(OutputFilename), False
    End If
    
    DeleteFile str & "temp1.bin"
    DeleteFile str & "temp2.bin"
    DeleteFile str & "temp3.bin"
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBMap"
    Close #tempNum1
    Close #tempNum2
    Close #tempNum3
    Close #nInputFilenum
    DeleteFile str & "temp1.bin"
    DeleteFile str & "temp2.bin"
    DeleteFile str & "temp3.bin"
End Sub


Public Sub RLENoDither(InputFilenum As Integer, OutputFilenum As Integer, HeaderSize As Integer, OutputFilename As String)

    On Error GoTo HandleErrors

    ReDim SpanArray(0) As tSpan

    'Read spans from output file
    CreateSpanTable OutputFilenum, HeaderSize, SpanArray
    
    'DebugPrint "c:\windows\desktop\orgdebug.txt", SpanArray
    
    Dim i As Integer
    Dim j As Integer
    
    Dim nCount As Integer

    'loop through spans
    For i = 0 To (UBound(SpanArray) - 1)
        
        If i > (UBound(SpanArray) - 1) Then
            Exit For
        End If
        
        If SpanArray(i).Length = 1 Then
            If SpanArray(i + 1).Length = 2 Then
          
                'add next 1 span to current
                ReDim Preserve SpanArray(i).iBytes(3)
                SpanArray(i).iBytes(1) = SpanArray(i + 1).iBytes(0)
                SpanArray(i).iBytes(2) = SpanArray(i + 1).iBytes(0)

                SpanArray(i).Length = 3
                
                'get rid of extra span from array
                For j = (i + 2) To UBound(SpanArray)
                    SpanArray(j - 1) = SpanArray(j)
                Next j
                
                If (i + 2) < UBound(SpanArray) Then
                    ReDim Preserve SpanArray(UBound(SpanArray) - 1)
                    nCount = nCount + 1
                End If

            End If
        End If
        
        nCount = nCount + 1
        
    Next i
    
    'remove redudant norm spans
    RemoveDoubleNorms SpanArray
    
    'DebugPrint "c:\windows\desktop\optdebug.txt", SpanArray
     
    'output new file using InputFilenum based on SpanArray
    OutputOptimizeFile InputFilenum, OutputFilenum, HeaderSize, SpanArray, OutputFilename

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modRLE:RLENoDither Error"
End Sub

