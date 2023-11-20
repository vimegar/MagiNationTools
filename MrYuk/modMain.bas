Attribute VB_Name = "modMain"
'***************************************************************************
'   VB Interface Setup
'***************************************************************************
    
    Option Explicit

'***************************************************************************
'   Global variables
'***************************************************************************

    Public glMainHDC As Long
    Public gbClosingApp As Boolean
    Public gbRefOpen As Boolean
    Public gnTool As Integer
    Public gsCurPath As String
    
    Public gResourceCache As New clsResourceCache
    Public gSelection As New clsSelection
    Public gPaletteSrcForm As Form

    Public gnWaitForPal As Integer
    Public gnPalCopy As Integer
    Public gnPalRed As Integer
    Public gnPalGreen As Integer
    Public gnPalBlue As Integer
    
    Public gGBCollisionCodes As New clsGBCollisionCodes
    Public gFrmColCodes As New frmEditCollisionCodes
    
    Public gBufferMap As New clsGBMap
    Public gBufferBackground As New clsGBBackground

    Public gsProjectPath As String
    Public gsPackToBinPath As String
    
    Public gbBatchGoing As Boolean
    
    Public gnReplaceTile As Integer
    
    Public gbBatchExport As Boolean
    Public gbBatchExportVRAMLoad As Boolean
    Public gbBatchExportColLoad As Boolean
    
    Public gbOpeningChild As Boolean
    Public gbAbort As Boolean
Public Function GetRecentFilename(ByVal sFilename As String) As String

    GetRecentFilename = "...\" & GetTruncFilename(sFilename)
    
End Function

Public Function CollectionFind(col As Collection, obj As Object) As Integer
    
    Dim cursor As Object
    Dim i As Integer

    For Each cursor In col
        i = i + 1
        If cursor Is obj Then
            CollectionFind = i
            Exit Function
        End If
    Next cursor
    
    CollectionFind = 0
    
End Function





Public Sub CreateDir(sDirname As String)

    On Error Resume Next
    MkDir sDirname
    
End Sub

Public Function GetIniData(ApplicationName As String, KeyName As String, IniFilename As String) As String
'********************************************
'
'   ApplicationName     Identification string
'   KeyName             Tag of the data to be retrieved
'   IniFilename         String containing the filename of
'                       the .ini file
'
'   This function returns the data stored within the ini
'   file that is specified.
'
'********************************************
        
    Dim dummy As String * 256
    
    GetPrivateProfileString ApplicationName, KeyName, "", dummy, 256, IniFilename
    GetIniData = Mid$(dummy, 1, InStr(dummy, Chr(0)) - 1)
    
End Function

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
            ret = ret + (d * val(Mid$(InputVal, i, 1)))
        End If
        nCount = nCount + 1
    Next i
    
    HexToDec = ret

End Function

Public Sub Main()

    Dim nMajor As Integer
    Dim nMinor As Integer
    Dim nRevision As Integer
    Dim appMajor As Integer
    Dim appMinor As Integer
    Dim appRev As Integer
    Dim nFilenum As Integer
    
    On Error Resume Next
    
    nFilenum = FreeFile
    
    Open "\\SERVER\MagiNation\GameBoy\Tools\Mr. Yuk\install.ini" For Input As #nFilenum
        Input #nFilenum, nMajor
        Input #nFilenum, nMinor
        Input #nFilenum, nRevision
    Close #nFilenum
  
    If App.Major = nMajor And App.Minor = nMinor And App.Revision = nRevision Then
    Else
        Dim ret As Integer
        ret = MsgBox("There is a new version (" & CStr(nMajor) & "." & CStr(nMinor) & "." & CStr(nRevision) & ") of Mr. Yuk on the server.  Please uninstall your current version and install the new one.  Do you still want to load Mr. Yuk?", vbYesNo Or vbQuestion, "New Version")
        If ret = vbNo Then
            End
        End If
    End If
 
    WriteIniData "MrYuk", "VersionMajor", App.Path & "\mryuk.ini", CStr(App.Major)
    WriteIniData "MrYuk", "VersionMinor", App.Path & "\mryuk.ini", CStr(App.Minor)
    WriteIniData "MrYuk", "VersionRevision", App.Path & "\mryuk.ini", CStr(App.Revision)

    Dim frm As New frmSplash
    Load frm
    frm.Left = (Screen.width - frm.width) / 2
    frm.Top = (Screen.height - frm.height) / 2
    AlwaysOnTop frm, True
    frm.Show
    
End Sub
Public Sub WriteIniData(ApplicationName As String, KeyName As String, IniFilename As String, DataString As String)
'********************************************
'
'   ApplicationName     Identification string
'   KeyName             Tag of the data to be retrieved
'   IniFilename         String containing the filename of
'                       the .ini file
'   DataString          The string to be written to the ini.
'
'   This procedure saves DataString under the tag KeyName
'   located in the ini file IniFilename.
'
'********************************************
        
    WritePrivateProfileString ApplicationName, KeyName, DataString, IniFilename
    
End Sub


Public Function Replace(ByVal StringToBeSearched As String, ByVal StringToFind As String, ByVal StringToReplaceWith) As String
'********************************************
'
'   StringToBeSearched      The overall string this will be
'                           checked for the target value.
'   StringToFind            The string that is being searched
'                           for in StringToBeSearched.
'   StringToReplaceWith     This is what all instances of
'                           StringToFind are replaced with.
'
'   This function returns a string that has had all instances
'   of StringToFind replaced with StringToReplaceWith.
'
'********************************************
        
    Dim loc As Long
    Dim firstChunk As String
    Dim lastChunk As String
    
    loc = InStr(StringToBeSearched, StringToFind)
    Do Until loc = 0
      
      firstChunk = Mid$(StringToBeSearched, 1, loc - 1)
      lastChunk = Mid$(StringToBeSearched, loc + Len(StringToFind))
      
      StringToBeSearched = firstChunk & StringToReplaceWith & lastChunk
      loc = InStr(StringToBeSearched, StringToFind)
      
    Loop
    
    Replace = StringToBeSearched

End Function

Public Function RoundUp(InputVal As Single) As Integer

    Dim d As Single
    
    d = (InputVal * 2) + 1
    
    RoundUp = d \ 2

End Function

Public Function NumToBitString(ByVal InputVal As Integer, Optional ByteLen As Integer) As String

    Dim i As Integer
    Dim ret As String
    
    If ByteLen = 0 Then
        ByteLen = 8
    End If
    
    For i = 0 To (ByteLen - 1)
      ret$ = ret$ & Format$(-(InputVal >= 2 ^ ((ByteLen - 1) - i)), "0")
      InputVal = InputVal - (2 ^ ((ByteLen - 1) - i) * -(InputVal >= 2 ^ ((ByteLen - 1) - i)))
    Next i
    NumToBitString = ret$

End Function

Public Function GetRGBFromLong(lColor As Long) As clsRGB

'***************************************************************************
'   Return a clsRGB from a long color value
'***************************************************************************

    Dim dummy As Long
    
    Set GetRGBFromLong = New clsRGB
    
    GetRGBFromLong.Red = lColor And &HFF
    GetRGBFromLong.Green = (Fix(lColor / 256)) And &HFF
    GetRGBFromLong.Blue = (Fix(lColor / 65536)) And &HFF
    
End Function

Public Function GetTruncFilename(ByVal sFilename As String) As String

'***************************************************************************
'   Returns only the filename portion of sFilename
'***************************************************************************

    Dim i As Integer
    
    For i = Len(sFilename) To 1 Step -1
        If Mid$(sFilename, i, 1) = "\" Then
            GetTruncFilename = Mid$(sFilename, i + 1)
            Exit Function
        End If
    Next i

    GetTruncFilename = sFilename

End Function

Public Sub AlwaysOnTop(Target As Form, TurnOn As Boolean)

'********************************************
'
'   Target      the target form.
'   TurnOn      a boolean value representing whether
'               or not the property is being turned
'               on or off.
'
'   This procedure sets the "always on top" property of
'   Target to the value of TurnOn.
'
'********************************************
    
    Dim ret As Long
    
    If TurnOn Then
      ret = SetWindowPos(Target.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H1 Or &H2)
    Else
      ret = SetWindowPos(Target.hwnd, -2, 0, 0, 0, 0, &H10 Or &H40 Or &H1 Or &H2)
    End If
    
End Sub

Public Function BitShiftRight(Operand As Variant, NumBits As Variant) As Integer

    BitShiftRight = CInt(Operand \ (2 ^ NumBits))

End Function

Public Function BitShiftLeft(Operand As Variant, NumBits As Variant) As Integer

    BitShiftLeft = CInt(Operand * (2 ^ NumBits))

End Function



Public Sub CleanUpForms(frm As Form)

    If frm Is Nothing Or frm.Name = "mdiMain" Then
        Exit Sub
    End If
       
    If (TypeOf frm Is frmEditBackground Or TypeOf frm Is frmEditCollisionCodes) And gbRefOpen Then
        frm.Left = -15000
        frm.Top = -15000
    Else
        
        Dim i As Integer
        Dim visForms As Integer
        
        For i = 0 To Forms.count - 1
            If Forms(i).Visible = True Then
                visForms = visForms + 1
            End If
        Next i
        
        frm.Left = (2 + (24 * (visForms - 1))) * Screen.TwipsPerPixelX
        frm.Top = (2 + (24 * (visForms - 1))) * Screen.TwipsPerPixelY
    End If

End Sub



Public Function KeyDown(KeyCode As Integer) As Boolean

    KeyDown = GetAsyncKeyState(CLng(KeyCode))

End Function

Public Sub SelectTool(tTool As GB_TOOLS)

    Dim btn As MSComctlLib.Button
    Set btn = mdiMain.ToolBar.Buttons(tTool + 1)
    
    mdiMain.ToolBar_ButtonClick btn
    btn.value = tbrPressed

End Sub



Public Function BitStringToNum(ByVal InputVal As String) As Integer

    Dim i As Integer
    Dim dummy As Integer
    
    For i = Len(InputVal$) To 1 Step -1
      dummy = dummy + (2 ^ (Len(InputVal$) - i) * -(val(Mid$(InputVal$, i, 1)) = "1"))
    Next i
    BitStringToNum = dummy

End Function

