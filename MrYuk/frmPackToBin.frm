VERSION 5.00
Begin VB.Form frmPackToBin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export To CGB File"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmPackToBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRLEBG 
      Caption         =   "&RLE BG"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbort 
      BackColor       =   &H000000FF&
      Caption         =   "ABORT!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox chkRLEBitmaps 
      Caption         =   "RLE &Bitmaps"
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   1800
      Width           =   1215
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Index           =   0
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.FileListBox File 
      Height          =   79260
      Index           =   0
      Left            =   5760
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame fraBatchExport 
      Caption         =   "&Batch Export"
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdBatch2 
         Caption         =   "&Standard <<"
         Height          =   360
         Left            =   3360
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose2 
         Caption         =   "&Close"
         Height          =   360
         Left            =   3360
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   21
         Top             =   675
         Width           =   2655
      End
      Begin VB.CommandButton cmdGetInput 
         Caption         =   "..."
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   675
         Width           =   300
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkSprites 
         Caption         =   "Sprites"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox chkVRAMs 
         Caption         =   "VRAMs"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkRawBG 
         Caption         =   "Raw BG"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkPatterns 
         Caption         =   "Patterns"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkPalettes 
         Caption         =   "Palettes"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkCollisionMaps 
         Caption         =   "Collision"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkBitmaps 
         Caption         =   "Bitmaps"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkMaps 
         Caption         =   "Maps"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdPickDir 
         Caption         =   "&Export To..."
         Height          =   360
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Current File:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label lblFile 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Input Folder:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   315
         Width           =   885
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Export Only These Types:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1125
         Width           =   1830
      End
   End
   Begin VB.Frame fraEditors 
      Caption         =   "Current &Editors:"
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdOpenBatch 
         Caption         =   "&Batch >>"
         Height          =   360
         Left            =   3360
         TabIndex        =   17
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdPack 
         Caption         =   "&Export"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   360
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   360
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   1200
      End
      Begin VB.ListBox lstEditors 
         Height          =   2595
         ItemData        =   "frmPackToBin.frx":030A
         Left            =   120
         List            =   "frmPackToBin.frx":030C
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmPackToBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 5010
Private Const DEF_HEIGHT = 3615

Private Type tEditor
    Filename As String
    Data As Object
End Type

Private mEditors() As tEditor
Private mbAuto As Boolean
Private msRecPath As String
Private mnRecCount As Integer

Private msInputPath As String
Private msOutputPath As String
Private msOutputFilename As String
Private msInputFilename As String

Private msFileList() As String
Private mbAbort As Boolean
Public Property Let bAuto(bNewValue As Boolean)

    mbAuto = bNewValue

End Property

Public Property Get bAuto() As Boolean

    bAuto = mbAuto

End Property

Public Sub mAddFileToList(Filename As String)

    msFileList(UBound(msFileList)) = Filename
    ReDim Preserve msFileList(UBound(msFileList) + 1)

End Sub
Private Sub mPackToRLEAfterBin(sFilename As String, HeaderSize As Integer)

    Screen.MousePointer = vbHourglass

    Dim i As Integer
    Dim str As String

    For i = Len(sFilename) To 1 Step -1
        If Mid$(sFilename, i, 1) = "\" Then
            str = Mid$(sFilename, 1, i)
            Exit For
        End If
    Next i

    PackRLE sFilename, str & "tempx.bin", HeaderSize
    DeleteFile sFilename
    CopyFile str & "tempx.bin", sFilename, False
    DeleteFile str & "tempx.bin"

End Sub

Private Sub mPopulateEditorList()

    On Error Resume Next
    
    Dim i As Integer
    Dim str As String
        
    ReDim mEditors(0)
        
    lstEditors.Clear
    
    For i = 0 To Forms.count - 1
            
        str = ""
        str = GetTruncFilename(Forms(i).sFilename)
        
        If str <> "" Then
        
            mEditors(UBound(mEditors)).Filename = str
        
            Select Case Forms(i).Name
                Case "frmEditBitmap"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBBitmap
                Case "frmEditMap"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBMap
                Case "frmEditPalette"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBPalette
                Case "frmEditBackground"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBBackground
                Case "frmEditVRAM"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBVRAM
                Case "frmEditSpriteGroup"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBSpriteGroup
                Case "frmEditCollisionMap"
                    Set mEditors(UBound(mEditors)).Data = Forms(i).GBCollisionMap
            End Select
        
            If Forms(i).Name <> "frmEditCollisionCodes" Then
                lstEditors.AddItem mEditors(UBound(mEditors)).Filename
                cmdPack.Enabled = True
                ReDim Preserve mEditors(UBound(mEditors) + 1)
            End If
            
        End If
    Next i

End Sub
Private Sub mSetDirPat()

    Dim str As String
    str = ""
    
    If chkBitmaps.value = vbChecked Then
        str = str & "BIT_*.BIT;"
    End If

    If chkCollisionMaps.value = vbChecked Then
        str = str & "COL_*.CLM;"
    End If

    If chkMaps.value = vbChecked Then
        str = str & "SCR_*.MAP;"
    End If

    If chkPatterns.value = vbChecked Then
        str = str & "PAT_*.PAT;"
    End If

    If chkPalettes.value = vbChecked Then
        str = str & "PAL_*.PAL;"
    End If

    If chkRawBG.value = vbChecked Then
        str = str & "BG_*.BG;"
    End If

    If chkVRAMs.value = vbChecked Then
        str = str & "VRM_*.VRM;"
    End If

    If chkSprites.value = vbChecked Then
        str = str & "SPR_*.SPR;"
    End If
    
    File(0).Pattern = Mid$(str, 1, Len(str) - 1)

End Sub

Private Sub chkAll_Click()

    Dim i As Integer
    
    For i = 0 To (Me.Controls.count - 1)
        If TypeOf Me.Controls(i) Is CheckBox Then
            If Me.Controls(i).Name <> "chkAll" And Me.Controls(i).Name <> "chkRLE" And Me.Controls(i).Name <> "chkRLEBitmaps" And Me.Controls(i).Name <> "chkRLEBG" Then
                Me.Controls(i).value = chkAll.value
                If chkAll.value = vbChecked Then
                    Me.Controls(i).Enabled = False
                Else
                    Me.Controls(i).Enabled = True
                End If
            End If
        End If
    Next i
    
End Sub


Private Sub cmdAbort_Click()

    mbAbort = True

End Sub

Private Sub cmdBatch2_Click()

    fraBatchExport.Visible = False

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub


Private Sub cmdClose2_Click()

    Unload Me

End Sub

Private Sub cmdGetInput_Click()

    msInputPath = UCase$(GetPathDialog)
    
    If msInputPath = "" Then
        msInputPath = txtInput.Text
        Exit Sub
    End If
    
    txtInput.Text = msInputPath

End Sub

Private Sub cmdOpenBatch_Click()

    fraBatchExport.Left = fraEditors.Left
    fraBatchExport.Top = fraEditors.Top

    fraBatchExport.Visible = True
    
End Sub

Public Sub cmdPack_Click()
 
    On Error GoTo HandleErrors

    If mbAbort Then
        cmdAbort.Visible = False
        Exit Sub
    End If

    Dim i As Integer
    Dim matt As String
    Dim nFilenum As Integer
    Dim HeaderSize As Integer
    
    Screen.MousePointer = vbHourglass
    
    If bAuto Then
        
        If Len(msInputFilename) > 40 Then
            Dim str As String
            str = Mid$(msInputFilename, Len(msInputFilename) - 37)
            
            For i = 1 To Len(str)
                
                DoEvents
                
                If Mid$(str, i, 1) = "\" Then
                    str = Mid$(str, i)
                    Exit For
                End If
            Next i
            
            lblFile.Caption = "..." & str
        Else
            lblFile.Caption = msInputFilename
        End If
        
        Dim sPack As String
        
        sPack = GetTruncFilename(msOutputFilename)
        
        For i = 1 To Len(sPack)
            If Mid$(sPack, i, 1) = "_" Then
                str = UCase$(Mid$(sPack, 1, i))
                Exit For
            End If
        Next i
        
        Dim cli As New intResourceClient
        
        HeaderSize = 0
        
        Dim pathBin As String
        Dim Paths As String
        Dim strBin As String
        Dim strS As String
        
        pathBin = msOutputPath & "\BIN\" & GetTruncFilename(msOutputFilename)
        Paths = msOutputPath & "\S\" & GetTruncFilename(msOutputFilename)
        
        CreateDir msOutputPath & "\BIN"
        CreateDir msOutputPath & "\S"
        
        For i = Len(pathBin) To 1 Step -1
            If Mid$(pathBin, i, 1) = "." Then
                strBin = Mid$(pathBin, 1, i - 1) & ".BIN"
                Exit For
            End If
        Next i
        
        For i = Len(Paths) To 1 Step -1
            If Mid$(Paths, i, 1) = "." Then
                strS = Mid$(Paths, 1, i - 1) & ".S"
                Exit For
            End If
        Next i
        
        Select Case str
            Case "VRM_"
                Dim oVRAM As clsGBVRAM
                Set oVRAM = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oVRAM Is Nothing Then
                    If Not oVRAM.CacheObject Is Nothing Then
                        oVRAM.PackToBin strS
                    End If
                End If
            Case "SPR_"
                Dim oSprite As clsGBSpriteGroup
                Set oSprite = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oSprite Is Nothing Then
                    If Not oSprite.CacheObject Is Nothing Then
                        oSprite.PackToBin strS
                    End If
                End If
            Case "BG_"
                Dim oBG As clsGBBackground
                Set oBG = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oBG Is Nothing Then
                    If Not oBG.CacheObject Is Nothing Then
                        
                        oBG.tBackgroundType = GB_RAWBG
                        
                        If chkRLEBG.value = vbChecked Then
                            oBG.PackToBinRLE strBin
                        Else
                            oBG.PackToBin strS
                        End If
                    End If
                End If
            Case "BIT_"
                Dim oBitmap As clsGBBitmap
                Set oBitmap = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oBitmap Is Nothing Then
                    If Not oBitmap.CacheObject Is Nothing Then
                        oBitmap.PackToBin strBin
                        If (chkRLEBitmaps.value = vbChecked) Then
                            HeaderSize = 0
                            mPackToRLEAfterBin strBin, HeaderSize
                        End If
                    End If
                End If
            Case "COL_"
                Dim oCollisionMap As clsGBCollisionMap
                Set oCollisionMap = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oCollisionMap Is Nothing Then
                    If Not oCollisionMap.CacheObject Is Nothing Then
                        oCollisionMap.PackToBin strBin
                        HeaderSize = 3
                        mPackToRLEAfterBin strBin, HeaderSize
                    End If
                End If
            Case "PAT_"
                Dim oPattern As clsGBBackground
                Set oPattern = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oPattern Is Nothing Then
                    If Not oPattern.CacheObject Is Nothing Then
                        oPattern.tBackgroundType = GB_PATTERNBG
                        oPattern.PackToBin strBin
                        mPackToRLEAfterBin strBin, HeaderSize
                    End If
                End If
            Case "PAL_"
                Dim oPalette As clsGBPalette
                Set oPalette = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oPalette Is Nothing Then
                    If Not oPalette.CacheObject Is Nothing Then
                        oPalette.PackToBin strBin
                    End If
                End If
            Case "SCR_"
                Dim oMap As clsGBMap
                Set oMap = gResourceCache.GetResourceFromFile(msInputFilename, cli)
                If Not oMap Is Nothing Then
                    If Not oMap.CacheObject Is Nothing Then
                        oMap.PackToBin strBin
                        HeaderSize = 7
                        mPackToRLEAfterBin strBin, HeaderSize
                    End If
                End If
        End Select
        
        gResourceCache.ReleaseClient cli
        
        Exit Sub
    End If
    
    For i = Len(mEditors(lstEditors.ListIndex).Filename) To 1 Step -1
        If Mid$(mEditors(lstEditors.ListIndex).Filename, i, 1) = "." Then
            matt = Mid$(mEditors(lstEditors.ListIndex).Filename, 1, i - 1)
            Exit For
        End If
    Next i
    
    If lstEditors.ListIndex < 0 Then
        MsgBox "You must select a valid editor from the list box!", vbInformation, "Information"
        Exit Sub
    End If
    
    If TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBSpriteGroup Or TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBVRAM Then
        
        With mdiMain.Dialog
    
            .InitDir = Mid$(gsPackToBinPath, 1, Len(gsPackToBinPath) - Len(GetTruncFilename(gsPackToBinPath)))
            .DefaultExt = "s"
            .DialogTitle = "Save to assembly file"
            .Filename = matt
            .Filter = "Assembly files (*.s)|*.s"
            .ShowSave
            If .Filename = matt Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            gsPackToBinPath = .Filename
        
            mEditors(lstEditors.ListIndex).Data.PackToBin .Filename
        
        End With
        
    ElseIf TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBBackground Then
        
        If mEditors(lstEditors.ListIndex).Data.tBackgroundType = GB_RAWBG Then
            
            With mdiMain.Dialog
        
                .InitDir = Mid$(gsPackToBinPath, 1, Len(gsPackToBinPath) - Len(GetTruncFilename(gsPackToBinPath)))
                .Filename = matt
                
                If chkRLEBG.value = vbChecked Then
                    .DefaultExt = "bin"
                    .DialogTitle = "Save to bin file"
                    .Filter = "Binary files (*.bin)|*.bin"
                Else
                    .DefaultExt = "s"
                    .DialogTitle = "Save to assembly file"
                    .Filter = "Assembly files (*.s)|*.s"
                End If
                
                .ShowSave
                If .Filename = matt Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                gsPackToBinPath = .Filename
            
                If chkRLEBG.value = vbChecked Then
                    mEditors(lstEditors.ListIndex).Data.PackToBinRLE .Filename
                Else
                    mEditors(lstEditors.ListIndex).Data.PackToBin .Filename
                End If
                
            End With
        Else
            GoTo lElse
        End If
    
    ElseIf TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBPalette Then
        
        With mdiMain.Dialog
    
            .InitDir = Mid$(gsPackToBinPath, 1, Len(gsPackToBinPath) - Len(GetTruncFilename(gsPackToBinPath)))
            .DefaultExt = "bin"
            .DialogTitle = "Save to bin file"
            .Filename = matt
            .Filter = "Bin files (*.bin)|*.bin"
            .ShowSave
            If .Filename = matt Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            gsPackToBinPath = .Filename
        
            mEditors(lstEditors.ListIndex).Data.PackToBin .Filename
            
        End With
        
    ElseIf TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBBitmap Then
        
        With mdiMain.Dialog
    
            .InitDir = Mid$(gsPackToBinPath, 1, Len(gsPackToBinPath) - Len(GetTruncFilename(gsPackToBinPath)))
            .DefaultExt = "bin"
            .DialogTitle = "Save to bin file"
            .Filename = matt
            .Filter = "Bin files (*.bin)|*.bin"
            .ShowSave
            If .Filename = matt Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            gsPackToBinPath = .Filename
        
            mEditors(lstEditors.ListIndex).Data.PackToBin .Filename
            
            If chkRLEBitmaps.value = vbChecked Then
                HeaderSize = 0
                mPackToRLEAfterBin .Filename, HeaderSize
            End If
            
        End With
        
    Else
lElse:
        With mdiMain.Dialog
    
            .InitDir = Mid$(gsPackToBinPath, 1, Len(gsPackToBinPath) - Len(GetTruncFilename(gsPackToBinPath)))
            .DefaultExt = "bin"
            .DialogTitle = "Save to bin file"
            .Filename = matt
            .Filter = "Bin files (*.bin)|*.bin"
            .ShowSave
            If .Filename = matt Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            gsPackToBinPath = .Filename
        
            mEditors(lstEditors.ListIndex).Data.PackToBin .Filename
            
            If TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBMap Then
                HeaderSize = 7
            ElseIf TypeOf mEditors(lstEditors.ListIndex).Data Is clsGBCollisionMap Then
                HeaderSize = 3
            Else
                HeaderSize = 0
            End If
    
            mPackToRLEAfterBin .Filename, HeaderSize

        End With
    
    End If
    
    MsgBox "Pack successfully completed!", vbInformation, "Success"
    Screen.MousePointer = 0
    
Exit Sub
                                 
HandleErrors:
    gResourceCache.ReleaseClient cli
    MsgBox "(" & Err.Description & ")" & vbCrLf & "An error has occured while loading a file!  This is most likely because the file being loaded is invalid.", vbCritical, "frmPackToBin:cmdPack_Click Error"
    Screen.MousePointer = 0
End Sub

Private Sub cmdPickDir_Click()

    On Error GoTo HandleErrors

    If txtInput.Text = "" Then
        MsgBox "You must choose an input path!", vbCritical, "Input Error"
        Exit Sub
    End If
    
    mSetDirPat

    msOutputPath = UCase$(GetPathDialog)

    If msOutputPath = "" Then
        Exit Sub
    End If

    If MsgBox("You must make sure the output folder is empty. If it is not, the output files may contain serious errors!  Do you still want to continue?", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    mbAbort = False
    cmdAbort.Visible = True
    
    msRecPath = ""
    mnRecCount = 0
    ReDim msFileList(0)
    ReDim msPriorityFiles(0)

    Screen.MousePointer = vbHourglass

    fraBatchExport.Enabled = False
    chkRLEBitmaps.Enabled = False
    chkRLEBG.Enabled = False

    gbBatchExport = True
    BatchPack msInputPath
    gbBatchExport = False
    
    Screen.MousePointer = 0
    
    fraBatchExport.Enabled = True
    chkRLEBitmaps.Enabled = True
    chkRLEBG.Enabled = True
    
    If mbAbort Then
        MsgBox "Batch export aborted!", vbCritical, "Abortion"
    Else
        MsgBox "Batch export completed successfully!", vbInformation, "Success"
    End If
    
    cmdAbort.Visible = False
    
    lblFile.Caption = ""

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmPackToBin:cmdPickDir_Click Error"
End Sub

Private Sub Form_Load()

    On Error GoTo HandleErrors

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT

    chkAll.value = vbChecked

    ReDim msFileList(0)

    mPopulateEditorList
    
    On Error Resume Next
    lstEditors.ListIndex = 0
    On Error GoTo HandleErrors
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmPackToBin:Form_Load Error"
End Sub

Public Function BatchPack(InputPath As String) As Boolean

    On Error GoTo HandleErrors

    If mbAbort Then
        cmdAbort.Visible = False
        Exit Function
    End If

    Dim dirCursor As Integer
    Dim fileCursor As Integer
    
    If InputPath = msRecPath Then
        Exit Function
    End If
    
    msRecPath = InputPath
    
    mnRecCount = mnRecCount + 1
    On Error Resume Next
    Load Folder(mnRecCount)
    Load File(mnRecCount)
    On Error GoTo HandleErrors
    Folder(mnRecCount).Path = InputPath
    
    If Folder(mnRecCount).ListCount > 0 Then
        Folder(mnRecCount).ListIndex = 0
    End If
    
    For dirCursor = Folder(mnRecCount).ListIndex To (Folder(mnRecCount).ListCount - 1)
        
        Dim bAdded As Boolean
        bAdded = BatchPack(Folder(mnRecCount).List(dirCursor))
        
        If Folder(mnRecCount).List(dirCursor) = "" Then
            GoTo bpEnd
        End If
        
        File(mnRecCount).Path = Folder(mnRecCount).List(dirCursor)
        
        For fileCursor = 0 To (File(mnRecCount).ListCount - 1)
        
            DoEvents
        
            If Not bAdded Then
                Dim sFilename As String
                sFilename = Folder(mnRecCount).List(dirCursor) & "\" & File(mnRecCount).List(fileCursor)
                
                msOutputFilename = UCase$(msOutputPath & "\" & GetTruncFilename(sFilename))
                msInputFilename = UCase$(msInputPath & Mid$(Folder(mnRecCount).List(dirCursor), Len(msInputPath) + 1) & "\" & GetTruncFilename(sFilename))
                
                Dim i As Integer
                Dim inputFile As String
                
                mAddFileToList msInputFilename
                inputFile = msFileList(UBound(msFileList) - 1)
                
                Dim inp As String
                inp = GetTruncFilename(inputFile)
                
                Dim bMatch As Boolean
                bMatch = False
                For i = 0 To (UBound(msFileList) - 2) 'doesnt include current entry
                    
                    If inp = GetTruncFilename(msFileList(i)) Then
                        
                        Dim srcPath As String
                        Dim k As Integer
                        For k = Len(inputFile) To 1 Step -1
                            If Mid$(inputFile, k, 1) = "\" Then
                                srcPath = Mid$(inputFile, 1, k)
                                Exit For
                            End If
                        Next k
                                
                        If UCase$(Dir(srcPath & inp & ".SOURCE")) = UCase$(inp & ".SOURCE") Then
                            
                            'overwrite file
                            msFileList(i) = inputFile
                            Dim str As String
                            Dim noextfile As String
                            Dim ext As String
                            For k = Len(msOutputFilename) To 1 Step -1
                                If Mid$(msOutputFilename, k, 1) = "\" Then
                                    str = Mid$(msOutputFilename, 1, k - 1)
                                    ext = Mid$(msOutputFilename, k)
                                    Exit For
                                End If
                            Next k
                            For k = Len(msOutputFilename) To 1 Step -1
                                If Mid$(msOutputFilename, k, 1) = "." Then
                                    noextfile = Mid$(msOutputFilename, 1, k - 1)
                                    Exit For
                                End If
                            Next k
                            
                            If UCase$(ext) = ".VRM" Or UCase$(ext) = ".BG" Or UCase$(ext) = ".SPR" Then
                                str = str & "\S\" & GetTruncFilename(noextfile) & ".S"
                            Else
                                str = str & "\BIN\" & GetTruncFilename(noextfile) & ".BIN"
                            End If
                            On Error Resume Next
                            Kill str
                            On Error GoTo HandleErrors
                                                               
                            mbAuto = True
                            cmdPack_Click
                            mbAuto = False
                                                               
                        End If
                        
                        ReDim Preserve msFileList(UBound(msFileList) - 1)
                        bMatch = True
                        Exit For
                        
                    End If
                Next i
                
                If Not bMatch Then
                    mbAuto = True
                    cmdPack_Click
                    mbAuto = False
                End If
                
            End If
            BatchPack = True
        Next fileCursor
    Next dirCursor

bpEnd:
    
    On Error Resume Next
    Unload Folder(mnRecCount)
    Unload File(mnRecCount)
    On Error GoTo 0
    mnRecCount = mnRecCount - 1

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmPackToBin:BatchPack Error"
End Function

Private Sub lstEditors_DblClick()

    cmdPack_Click

End Sub


