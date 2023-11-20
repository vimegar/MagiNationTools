Attribute VB_Name = "modFileIO"
Option Explicit

Public Function UnpackFile(sFilename As String) As intResource

'***************************************************************************
'   Create new GB bitmap and open it in an editor
'***************************************************************************

    If sFilename = "" Then
        Exit Function
    End If

'Setup error handling
    On Error GoTo HandleErrors
    
'Get file type from file
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Dim str As String
    Dim sdFilename As String
    'Dim d As Object
    Dim sFilter As String
    Dim sDir As String
    Dim flag As Boolean
    Dim sParentPath As String
    
    If InStr(sFilename, "\") Then
        Dim i As Integer
        For i = Len(sFilename) To 1 Step -1
            If Mid$(sFilename, i, 1) = "\" And flag Then
                sParentPath = Mid$(sFilename, 1, i - 1)
                sdFilename = sFilename
                GoTo ufOpen:
            End If
            If Mid$(sFilename, i, 1) = "\" Then
                flag = True
            End If
        Next i
    End If

ufOpen:
    
    sFilename = GetTruncFilename(sFilename)
    
    Select Case UCase$(Mid$(sFilename, InStr(sFilename, ".")))
        Case ".BIT"
            str = "BITMAP"
            'Set d = New clsGBBitmap
            sFilter = "GB Bitmaps (*.bit)|*.bit"
            sDir = "Bitmaps"
        Case ".VRM"
            str = "VRAM"
            'Set d = New clsGBVRAM
            sFilter = "GB VRAMs (*.vrm)|*.vrm"
            sDir = "VRAMs"
        Case ".PAT"
            str = "PATTERN"
            'Dim d2 As New clsGBBackground
            'd2.tBackgroundType = GB_PATTERNBG
            'Set d = d2
            sFilter = "GB Patterns (*.pat)|*.pat"
            sDir = "Patterns"
        Case ".MAP"
            str = "MAP"
            'Set d = New clsGBMap
            sFilter = "GB Maps (*.map)|*.map"
            sDir = "Maps"
        Case ".PAL"
            str = "PALETTE"
            'Set d = New clsGBPalette
            sFilter = "GB Palettes (*.pal)|*.pal"
            sDir = "Palettes"
        Case ".CLM"
            str = "COLLISION MAP"
            'Set d = New clsGBCollisionMap
            sFilter = "GB Collision Maps (*.clm)|*.clm"
            sDir = "Collision"
        Case ".BG"
            str = "BACKGROUND"
            'd2.tBackgroundType = GB_RAWBG
            'Set d = d2
            sFilter = "GB Backgrounds (*.bg)|*.bg"
            sDir = "Backgrounds"
        Case ".SPR"
            str = "SPRITE GROUP"
            'Set d = New clsGBSpriteGroup
            sFilter = "GB Sprite Groups (*.spr)|*.spr"
            sDir = "Sprites"
        Case ".CLC"
            GoTo ufOpen2
    End Select
    
    If UCase$(Dir(sParentPath & "\" & sDir & "\" & GetTruncFilename(sFilename))) <> UCase$(GetTruncFilename(sFilename)) Then
        
        If Not gbOpeningChild Then
            If (sDir = "Bitmaps") Or (sDir = "Palettes") Then
                GoTo ufOpen2
            End If
        End If
        
        If gbBatchExportColLoad Then
            Exit Function
        End If
                
        MsgBox "This file has a broken link to the " & str & " file, " & sFilename & ".  Please find the missing file.", vbInformation, "Question"
                
        With mdiMain.Dialog
            .InitDir = sParentPath & "\" & sDir
            .DialogTitle = "Locate " & GetTruncFilename(sFilename)
            .Filename = ""
            .Filter = sFilter
            .ShowOpen
            If .Filename = "" Then
                Set UnpackFile = Nothing
                Exit Function
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, sParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, sParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
            sdFilename = .Filename
        End With
    Else
        sdFilename = sParentPath & "\" & sDir & "\" & sFilename
    End If

ufOpen2:
    
    Name sdFilename As UCase$(sdFilename)
    
    Open sdFilename For Binary Access Read Write As #nFilenum
    
    Dim fileType As Byte
    Get #nFilenum, , fileType
    
'Unpack file
    Dim NewInterface As New intResource
      
    Select Case fileType
        Case GB_BITMAP
            Set NewInterface = New clsGBBitmap
        Case GB_VRAM
            Set NewInterface = New clsGBVRAM
        Case GB_PALETTE
            Set NewInterface = New clsGBPalette
        Case GB_PATTERN
            Dim od As New clsGBBackground
            od.tBackgroundType = GB_PATTERNBG
            Set NewInterface = od
        Case GB_BG
            od.tBackgroundType = GB_RAWBG
            Set NewInterface = od
        Case GB_MAP
            Set NewInterface = New clsGBMap
        Case GB_COLLISIONCODES
            Set NewInterface = New clsGBCollisionCodes
        Case GB_COLLISIONMAP
            Set NewInterface = New clsGBCollisionMap
        Case GB_SPRITEGROUP
            Set NewInterface = New clsGBSpriteGroup
        Case Else
            Set UnpackFile = Nothing
            Exit Function
    End Select
    
    NewInterface.ParentPath = sParentPath
    
    Dim ret As Boolean
    ret = NewInterface.Unpack(nFilenum)

    If ret = False Then
        Set UnpackFile = Nothing
        Close #nFilenum
        Exit Function
    End If

    PackFile sdFilename, NewInterface

    Close #nFilenum
    
'Return appropriate object
    Set UnpackFile = NewInterface
    
'***************************************************************************

Exit Function

HandleErrors:
    Close #nFilenum
    MsgBox Err.Description, vbCritical, "modFileIO:UnpackFile Error"
End Function

Public Sub PackFile(ByVal sFilename As String, iResource As intResource)

    On Error GoTo HandleErrors

    If iResource Is Nothing Or sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile

    Open sFilename For Binary Access Write As #nFilenum
        iResource.Pack nFilenum
    Close #nFilenum
    
Exit Sub
                                 
HandleErrors:
    If Err.Description = "Path not found" Then
        MsgBox "(" & Err.Description & ")" & vbCrLf & "The file was not saved!  It appears that the file " & UCase$(GetTruncFilename(sFilename)) & " is being saved to a folder that is not part of a Mr. Yuk project structure.  Please make sure you have created a Mr. Yuk project to store your files!", vbInformation, "Path Error"
    Else
        MsgBox Err.Description, vbCritical, "modFileIO:PackFile Error"
    End If
End Sub
