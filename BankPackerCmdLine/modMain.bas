Attribute VB_Name = "modMain"
Option Explicit

Private Type tBank
    Size As Long
    Files() As String
'    Paths() As String
End Type

Private msBinFiles() As String
'Private msBinPaths() As String
Private mtBanks() As tBank
Private mnStartBank As Integer
Private mnEndBank As Integer
Private msFilename As String

Public Function GetTruncFilename(ByVal msFilename As String) As String

'***************************************************************************
'   Returns only the filename portion of msFilename
'***************************************************************************

    Dim i As Integer
    
    For i = Len(msFilename) To 1 Step -1
        If Mid$(msFilename, i, 1) = "\" Then
            GetTruncFilename = Mid$(msFilename, i + 1)
            Exit Function
        End If
    Next i

    GetTruncFilename = msFilename

End Function

Public Sub Main()
    
    On Error GoTo HandleErrors

    Dim i As Integer
    Dim cmdLine As String
    Dim curPos As Integer
    
    ReDim msBinFiles(0)
    
    cmdLine = Command
    
    msFilename = Mid$(cmdLine, 1, InStr(cmdLine, " "))
    curPos = Len(msFilename) + 1
    
    mnStartBank = HexToDec(Mid$(cmdLine, curPos, 2))
    curPos = curPos + 2
    
    mnEndBank = HexToDec(Mid$(cmdLine, curPos, 2))
    curPos = curPos + 2
    
    For i = curPos To Len(cmdLine)
        If Mid$(cmdLine, i, 1) = " " Then
            curPos = i + 1
            Exit For
        End If
    Next i
    
    Do Until curPos >= Len(cmdLine)
        For i = curPos To Len(cmdLine)
            If Mid$(cmdLine, i, 1) = " " Or i = Len(cmdLine) Then
                ReDim Preserve msBinFiles(UBound(msBinFiles) + 1)
                If i = Len(cmdLine) Then
                    msBinFiles(UBound(msBinFiles)) = Mid$(cmdLine, curPos)
                Else
                    msBinFiles(UBound(msBinFiles)) = Mid$(cmdLine, curPos, i - curPos)
                End If
                curPos = i + 1
                Exit For
            End If
        Next i
    Loop
    
    PackBanks
    Output

    End

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "BankPacker Error"
End Sub

Public Function HexToDec(InputVal As String) As Long

    Dim i As Integer
    Dim ret As Long
    Dim d As Long
    Dim nCount As Integer
    
    For i = Len(InputVal) To 1 Step -1
        
        d = 16 ^ nCount
        
        If InStr("ABCDEFabcdef", Mid$(InputVal, i, 1)) Then
            ret = ret + (d * (Asc(UCase$(Mid$(InputVal, i, 1))) - 55))
        Else
            ret = ret + (d * Val(Mid$(InputVal, i, 1)))
        End If
        nCount = nCount + 1
    Next i
    
    HexToDec = ret

End Function

Public Sub Output()

    Dim sPath As String
    Dim i As Integer
    Dim j As Integer
    Dim strTab As String
    Dim strBankNum As String
    Dim strDPath As String
    Dim dummy As String
    
    strTab = Chr(vbKeyTab)
    
    Dim nFilenum As String
    nFilenum = FreeFile
    
    Open CurDir & "\" & msFilename For Output As #nFilenum

    Print #nFilenum, ";********************************"
    Print #nFilenum, "; " & UCase$(GetTruncFilename(msFilename))
    Print #nFilenum, ";********************************"
    Print #nFilenum, ";" & strTab & "Author:" & strTab & "Mr. Yuck"
    Print #nFilenum, ";" & strTab & "(c)2000" & strTab & "Interactive Imagination"
    Print #nFilenum, ";" & strTab & "All rights reserved"
    Print #nFilenum, ""
    Print #nFilenum, ";********************************"
    Print #nFilenum, strTab & strTab & strTab & strTab & "LIB" & strTab & strTab & "GLOBALS.S"
    Print #nFilenum, ""
    
    For i = mnStartBank To mnEndBank
        
        strBankNum = Hex(i)
        If Len(strBankNum) < 2 Then
            strBankNum = "0" & strBankNum
        End If

        Print #nFilenum, "BANK" & strBankNum & strTab & strTab & strTab & "GROUP" & strTab & "$" & strBankNum
        Print #nFilenum, strTab & strTab & strTab & strTab & "ORG" & strTab & strTab & "$4000"
        Print #nFilenum, ""
        
        For j = 1 To UBound(mtBanks(i).Files)
            dummy = UCase$(Mid$(GetTruncFilename(mtBanks(i).Files(j)), 1, Len(GetTruncFilename(mtBanks(i).Files(j))) - 4))
            Print #nFilenum, dummy & String((8 - Len(dummy)) * -(Len(dummy) <= 8), " ") & strTab & "LIBBIN" & strTab & UCase$(GetTruncFilename(mtBanks(i).Files(j)))
        Next j
        
        Print #nFilenum, ""
        
    Next i
    
    Print #nFilenum, ""
    Print #nFilenum, ";********************************"
    Print #nFilenum, strTab & "END"
    Print #nFilenum, ";********************************"
        
    Close #nFilenum

End Sub


Public Sub PackBanks()

    On Error GoTo HandleErrors
    
    Dim i As Integer
    Dim bigFile As Integer
    Dim fileSize As Long
    Dim curBank As Integer
    Dim sUsedFiles As String
    Dim nCount As Integer
    
    ReDim mtBanks(128)
    
    For i = 0 To 127
        ReDim mtBanks(i).Files(0)
'        ReDim mtBanks(i).Paths(0)
    Next i
    
    Do Until nCount >= UBound(msBinFiles)
        bigFile = 0
        
        For i = 1 To UBound(msBinFiles)
            fileSize = FileLen(msBinFiles(i))
            
            If bigFile = 0 And InStr(sUsedFiles, Format(CStr(i), "000") & ";") = 0 Then
                bigFile = i
            Else
                If Dir(msBinFiles(bigFile)) = GetTruncFilename(msBinFiles(bigFile)) Then
                    If (fileSize > FileLen(msBinFiles(bigFile))) And InStr(sUsedFiles, Format(CStr(i), "000") & ";") = 0 Then
                        bigFile = i
                    End If
                End If
            End If
        Next i
        
        nCount = nCount + 1
        sUsedFiles = sUsedFiles & Format(CStr(bigFile), "000") & ";"
    
        'get curBank
        For i = mnStartBank To mnEndBank
            If mtBanks(i).Size + FileLen(msBinFiles(bigFile)) <= &H4000 Then
                curBank = i
                Exit For
            End If
        Next i
    
        'pack
        ReDim Preserve mtBanks(curBank).Files(UBound(mtBanks(curBank).Files) + 1)
'        ReDim Preserve mtBanks(curBank).Paths(UBound(mtBanks(curBank).Paths) + 1)
        
        With mtBanks(curBank)
            .Files(UBound(.Files)) = msBinFiles(bigFile)
'            .Paths(UBound(.Paths)) = msBinPaths(bigFile)
            .Size = .Size + FileLen(msBinFiles(bigFile))
        End With
    
    Loop
Exit Sub

HandleErrors:
    If Err.Description = "File not found" Then
        MsgBox Err.Description & ":" & vbCrLf & msBinFiles(i), vbCritical, "BankPacker Error"
    Else
        MsgBox Err.Description, vbCritical, "BankPacker Error"
    End If
End Sub
