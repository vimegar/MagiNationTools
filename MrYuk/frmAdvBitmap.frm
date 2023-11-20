VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdvBitmap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Bitmap Tool"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmAdvBitmap.frx":0000
   LinkTopic       =   "frmEditAdvBitmap"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   Begin VB.CheckBox chkBitmapOnly 
      Caption         =   "&Bitmap Only"
      Height          =   255
      Left            =   1320
      TabIndex        =   33
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox picVRAMBitmap 
      AutoRedraw      =   -1  'True
      Height          =   3840
      Left            =   7560
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Frame fraBatch 
      Caption         =   "B&atch Operation"
      Enabled         =   0   'False
      Height          =   5295
      Left            =   3360
      TabIndex        =   25
      Top             =   30
      Width           =   3015
      Begin VB.CheckBox chkPackToGB 
         Caption         =   "Automatically Export to CGB"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   840
         Width           =   2295
      End
      Begin VB.FileListBox FileList 
         Height          =   3795
         Left            =   120
         Pattern         =   "*.qwewq"
         TabIndex        =   29
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtInputDir 
         Enabled         =   0   'False
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
         Left            =   465
         TabIndex        =   27
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdBrowseInput 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   26
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Input Folder:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   885
      End
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   345
      Left            =   3900
      TabIndex        =   21
      Top             =   5415
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Max             =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   5385
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6773
            MinWidth        =   6773
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOutput 
      Caption         =   "&Output"
      Height          =   5295
      Left            =   120
      TabIndex        =   19
      Top             =   30
      Width           =   3135
      Begin VB.TextBox txtEnd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   35
         Text            =   "7"
         Top             =   1200
         Width           =   330
      End
      Begin VB.TextBox txtStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   34
         Text            =   "0"
         Top             =   840
         Width           =   330
      End
      Begin VB.TextBox txtVrmPrefix 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "VRM_"
         Top             =   1560
         Width           =   1290
      End
      Begin VB.CommandButton cmdBrowseVRAM 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox txtVRAMFile 
         Enabled         =   0   'False
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
         Left            =   765
         TabIndex        =   7
         Top             =   2640
         Width           =   2205
      End
      Begin VB.CheckBox chkVRAM 
         Caption         =   "&VRAM"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Frame fraVRAM 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton optVRAMType 
            Caption         =   "Use E&xisiting VRAM File"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   525
            Width           =   2415
         End
         Begin VB.OptionButton optVRAMType 
            Caption         =   "&Create"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   165
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.OptionButton optBGType 
         Caption         =   "Pat&tern"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   3555
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optBGType 
         Caption         =   "Raw &BG"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "C&lose"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   4830
         Width           =   1290
      End
      Begin VB.CheckBox chkBatch 
         Caption         =   "&Batch"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   3480
         Width           =   855
      End
      Begin VB.CheckBox chkOpenFiles 
         Caption         =   "&Open Editors"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   3780
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   4455
         Width           =   255
      End
      Begin VB.TextBox txtOutputDir 
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
         Left            =   540
         TabIndex        =   15
         Text            =   "C:\WINDOWS\DESKTOP\"
         Top             =   4440
         Width           =   2430
      End
      Begin VB.TextBox txtPatPrefix 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Text            =   "PAT_"
         Top             =   3120
         Width           =   1290
      End
      Begin VB.TextBox txtPalPrefix 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "PAL_"
         Top             =   480
         Width           =   1290
      End
      Begin VB.CommandButton cmdCreateFiles 
         Caption         =   "&Create Files..."
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   4830
         Width           =   1335
      End
      Begin VB.CheckBox chkPalette 
         Caption         =   "&Palette"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkPattern 
         Caption         =   "&Pattern"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "End:"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   37
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Start:"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   36
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Output &Folder:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   4170
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Include:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Prefixes:"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.PictureBox picBitmap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   7560
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "frmAdvBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 6525
Private Const DEF_HEIGHT = 6135




Private Sub mCreateResources(sBitmapFilename As String, sVRAMFilename As String)

    On Error GoTo HandleErrors
    
    TileGetterErrorCount = 0
    
    Screen.MousePointer = vbHourglass
    
    If txtOutputDir.Text = "" Then
        MsgBox "You must enter an output folder!", vbInformation, "Input Error"
        GoTo rExit
    Else
        If Mid$(txtOutputDir.Text, Len(txtOutputDir.Text), 1) <> "\" Then
            txtOutputDir.Text = txtOutputDir.Text & "\"
        End If
    End If
    
    Dim oCTArrayPat As New clsCTArray
    Dim oCTArrayVrm As New clsCTArray
    Dim oVRAM As New clsGBVRAM
    Dim oPat As New clsGBBackground
    Dim oBit As New clsGBBitmap
    Dim oVRAMBit As New clsGBBitmap
    Dim frmVrm As New frmEditVRAM
    Dim frmPal As New frmEditPalette
    Dim frmPat As New frmEditBackground
    Dim frmBit As New frmEditBitmap
    
    Dim BMPOffscreen As New clsOffscreen
    BMPOffscreen.CreateBitmapFromBMP sBitmapFilename
    
    If sVRAMFilename = "" Then
        CreateCTArrayFromBMP sBitmapFilename, BMPOffscreen, oBit, oCTArrayPat, Progress, Status
    Else
        Set oVRAM = gResourceCache.GetResourceFromFile(sVRAMFilename, frmVrm)
        CreateCTArrayFromBMP sBitmapFilename, BMPOffscreen, oBit, oCTArrayPat, Progress, Status
    End If
        
    If optBGType(0).value = True Then
        oPat.tBackgroundType = GB_PATTERNBG
        oPat.nWidth = 32
        oPat.nHeight = 32
    Else
        oPat.tBackgroundType = GB_RAWBG
        oPat.nWidth = oBit.width \ 8
        oPat.nHeight = oBit.height \ 8
    End If
        
    If sVRAMFilename = "" Then
        CreatePatternAndVRAM oCTArrayPat, oBit, oVRAM, oPat, oVRAMBit
    Else
        CreateCTArrayFromVRAM oCTArrayVrm, oVRAM
        CreatePattern oCTArrayPat, oCTArrayVrm, oPat
    End If
    
    Dim str As String
    Dim sFileWord As String
    Dim sWorkingDir As String
    
    str = GetTruncFilename(sBitmapFilename)
    sFileWord = Mid$(str, 1, Len(str) - 4)
    
    If chkBatch.value = vbChecked Then
        sWorkingDir = txtOutputDir.Text & "\"
        gbBatchGoing = True
    Else
        sWorkingDir = txtOutputDir.Text
    End If
    
    str = sWorkingDir
    
    CreateDir str
    CreateDir str & "\Bitmaps"
    CreateDir str & "\Collision"
    CreateDir str & "\Maps"
    CreateDir str & "\Palettes"
    CreateDir str & "\Patterns"
    CreateDir str & "\VRAMs"
    CreateDir str & "\Sprites"
    CreateDir str & "\Backgrounds"
    CreateDir str & "\Output"
    CreateDir str & "\Output\Bitmaps"
    CreateDir str & "\Output\Collision"
    CreateDir str & "\Output\Maps"
    CreateDir str & "\Output\Palettes"
    CreateDir str & "\Output\Patterns"
    CreateDir str & "\Output\VRAMs"
    CreateDir str & "\Output\Sprites"
    CreateDir str & "\Output\Backgrounds"

    Dim i As Integer
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    For i = Len(sWorkingDir) - 1 To 1 Step -1
        If Mid$(sWorkingDir, i, 1) = "\" Then
            str = Mid$(sWorkingDir, 1, i - 1)
            Exit For
        End If
    Next i
            
    Open str & "\dimensions.txt" For Append As #nFilenum
        Print #nFilenum, "[" & sFileWord & " Size]"
        Print #nFilenum, oBit.width \ 8
        Print #nFilenum, oBit.height \ 8
        Print #nFilenum, ""
    Close #nFilenum

    If chkBitmapOnly.value = vbChecked Then
        oBit.PackToBin sWorkingDir & "Output\Bitmaps\BIT_" & sFileWord & ".bin"
        GoTo rExit
    End If

    If chkPalette.value = vbChecked Then
        
        Set frmPal.GBPalette = oBit.GBPalette
        frmPal.sFilename = sWorkingDir & "Palettes\" & txtPalPrefix.Text & sFileWord & ".pal"
        
        gResourceCache.AddResourceToCache frmPal.sFilename, frmPal.GBPalette, frmPal
        
        frmPal.GBPalette.iStart = val(txtStart.Text)
        frmPal.GBPalette.iEnd = val(txtEnd.Text)
        
        If chkBatch.value = vbChecked Then
            PackFile frmPal.sFilename, oBit.GBPalette
            frmPal.bChanged = False
            
            If chkPackToGB.value = vbChecked Then
                str = GetTruncFilename(frmPal.sFilename)
                oBit.GBPalette.PackToBin sWorkingDir & "Output\Palettes\" & Mid$(str, 1, InStr(str, ".") - 1) & ".bin"
            End If
            
        Else
            frmPal.bChanged = True
        End If
        
        If chkOpenFiles.value = vbChecked Then
            frmPal.Show
        Else
            gResourceCache.ReleaseClient frmPal
            Unload frmPal
        End If
        
    End If

    If optVRAMType(1).value = True Then
    
        Set frmBit.GBBitmap = oVRAMBit
        
        frmBit.sFilename = sWorkingDir & "Bitmaps\BIT_" & sFileWord & ".bit"
       
        gResourceCache.AddResourceToCache frmBit.sFilename, frmBit.GBBitmap, frmBit
        
        If chkBatch.value = vbChecked Then
            PackFile frmBit.sFilename, oVRAMBit
            frmBit.bChanged = False
            
            If chkPackToGB.value = vbChecked Then
                str = GetTruncFilename(frmBit.sFilename)
                oVRAMBit.PackToBin sWorkingDir & "Output\Bitmaps\" & Mid$(str, 1, InStr(str, ".") - 1) & ".bin"
            End If
            
        Else
            frmBit.bChanged = True
        End If
        
        If chkOpenFiles.value = vbChecked Then
            frmBit.Show
        Else
            gResourceCache.ReleaseClient frmBit
            Unload frmBit
        End If
    Else
        gResourceCache.ReleaseClient frmBit
        Unload frmBit
    End If
        
    If chkVRAM.value = vbChecked Then
    
        Set frmVrm.GBVRAM = oVRAM
        
        If sVRAMFilename = "" Then
            frmVrm.sFilename = sWorkingDir & "VRAMs\" & txtVrmPrefix.Text & sFileWord & ".vrm"
        Else
            frmVrm.sFilename = sVRAMFilename
        End If
        
        If optVRAMType(1).value = True Then
            frmVrm.GBVRAM.AddVRAMEntry frmBit.sFilename, 36864, 0
        End If
        
        gResourceCache.AddResourceToCache frmVrm.sFilename, frmVrm.GBVRAM, frmVrm
        
        If chkBatch.value = vbChecked Then
            PackFile sWorkingDir & "VRAMs\" & txtVrmPrefix.Text & sFileWord & ".vrm", oVRAM
            frmVrm.bChanged = False
            
            If chkPackToGB.value = vbChecked Then
                str = GetTruncFilename(sWorkingDir & "VRAMs\" & txtVrmPrefix.Text & sFileWord & ".vrm")
                oVRAM.PackToBin sWorkingDir & "Output\VRAMs\" & Mid$(str, 1, InStr(str, ".") - 1) & ".bin"
            End If
            
        Else
            frmVrm.bChanged = True
        End If
        
        If chkOpenFiles.value = vbChecked Then
            frmVrm.Show
        Else
            gResourceCache.ReleaseClient frmVrm
            Unload frmVrm
        End If
    End If
    
    If chkPattern.value = vbChecked Then
            
        Set oPat.GBPalette = oBit.GBPalette
        Set oPat.GBVRAM = oVRAM
        oPat.sPaletteFile = sWorkingDir & "Palettes\" & txtPalPrefix.Text & sFileWord & ".pal"
        
        oPat.sVRAMFile = frmVrm.sFilename
                    
        If optBGType(0).value Then
            frmPat.sFilename = sWorkingDir & "Patterns\" & txtPatPrefix.Text & sFileWord & ".pat"
        Else
            frmPat.sFilename = sWorkingDir & "Backgrounds\BG_" & sFileWord & ".bg"
        End If
        
        frmPat.bOpening = True
        Set frmPat.GBBackground = oPat
        
        gResourceCache.AddResourceToCache frmPat.sFilename, frmPat.GBBackground, frmPat
        
        If chkBatch.value = vbChecked Then
            PackFile frmPat.sFilename, frmPat.GBBackground
            frmPat.bChanged = False
        
            str = GetTruncFilename(frmPat.sFilename)
            
            If chkPackToGB.value = vbChecked Then
                If optBGType(0).value Then
                    frmPat.GBBackground.PackToBin sWorkingDir & "Output\Patterns\" & Mid$(str, 1, InStr(str, ".") - 1) & ".s"
                Else
                    frmPat.GBBackground.PackToBin sWorkingDir & "Output\Backgrounds\" & Mid$(str, 1, InStr(str, ".") - 1) & ".s"
                End If
            End If
        
        Else
            frmPat.bChanged = True
        End If
        
        If chkOpenFiles.value = vbChecked Then
            frmPat.Show
        Else
            gResourceCache.ReleaseClient frmPat
            frmPat.bChanged = False
            Unload frmPat
        End If
               
    End If

    If TileGetterErrorCount > 0 Then
        Dim frm As New frmTileGetterErrors
        frm.sName = sFileWord
        frm.sOutputFolder = txtOutputDir.Text
        Set frm.Offscreen = BMPOffscreen
        Load frm
    End If

rExit:
    
    BMPOffscreen.Delete
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmAdvBitmap:mCreateResources Error"
    GoTo rExit
End Sub

Private Sub mMatchPal(oBitmap As clsGBBitmap, iPalMap() As Byte)

    Dim bitColor As New clsRGB
    Dim i As Integer
    Dim xTile As Integer
    Dim yTile As Integer
    Dim xPixel As Byte
    Dim yPixel As Byte
    Dim pal As Byte
    Dim colr As Byte
    Dim highPalMatch(7) As Byte
    Dim highPal As Byte
    Dim palMatch As Byte
    
    ReDim iPalMap((oBitmap.Offscreen.width \ 8), (oBitmap.Offscreen.height \ 8))
    
    Status.Panels(1).Text = "Matching palettes..."
    Progress.value = 0
    Progress.Max = (oBitmap.width \ 8) * (oBitmap.height \ 8)
    
    BitBlt picBitmap.hdc, 0, 0, picBitmap.Picture.width, picBitmap.Picture.height - 8, picBitmap.hdc, 0, 8, vbSrcCopy
    
    For yTile = 0 To (oBitmap.Offscreen.height \ 8) - 1
        For xTile = 0 To (oBitmap.Offscreen.width \ 8) - 1
            
            For pal = 0 To 7
                For yPixel = 0 To 7
                    For xPixel = 0 To 7
                        For colr = 1 To 4
                            
                            DoEvents
                            
                            'get color of pixel
                            Set bitColor = GetRGBFromLong(GetPixel(picBitmap.hdc, xPixel + (xTile * 8), yPixel + (yTile * 8)))
                            With bitColor
                                .Red = (.Red \ 8) And &H1F
                                .Green = (.Green \ 8) And &H1F
                                .Blue = (.Blue \ 8) And &H1F
                            End With
                            
                            'does pixel match?
                            If bitColor.Red = oBitmap.GBPalette.Colors(colr + (pal * 4)).Red And bitColor.Green = oBitmap.GBPalette.Colors(colr + (pal * 4)).Green And bitColor.Blue = oBitmap.GBPalette.Colors(colr + (pal * 4)).Blue Then
                                palMatch = palMatch + 1
                                GoTo NextPixel
                            End If
                            
                        Next colr
NextPixel:
                    Next xPixel
                Next yPixel
                
                If palMatch > highPalMatch(highPal) Then
                    highPalMatch(pal) = palMatch
                    highPal = pal
                End If
                palMatch = 0
                
            Next pal
        
            iPalMap(xTile, yTile) = highPal
            highPal = 0
            For i = 0 To 7
                highPalMatch(i) = 0
            Next i
        
            Progress.value = Progress.value + 1
        
        Next xTile
    Next yTile

    mRenderBitmap oBitmap, iPalMap

End Sub









Private Sub mRenderBitmap(oBitmap As clsGBBitmap, iPalMap() As Byte)

    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim PixelX As Byte
    Dim PixelY As Byte
    Dim pixelRGB As New clsRGB
    Dim PalID As Byte
    Dim bestColor As Byte

    Progress.value = 0
    Progress.Max = (oBitmap.height \ 8) * (oBitmap.width \ 8)
    Status.Panels(1).Text = "Rendering bitmap..."

    For Y = 0 To (oBitmap.Offscreen.height \ 8) - 1
        For X = 0 To (oBitmap.Offscreen.width \ 8) - 1
            For PixelY = 0 To 7
                For PixelX = 0 To 7

                    Set pixelRGB = GetRGBFromLong(GetPixel(picBitmap.hdc, (X * 8) + PixelX, (Y * 8) + PixelY))
                    With pixelRGB
                        .Red = (.Red \ 8) And &H1F
                        .Green = (.Green \ 8) And &H1F
                        .Blue = (.Blue \ 8) And &H1F
                    End With
                    PalID = iPalMap(X, Y)

                    bestColor = 0
                    For i = 1 To 4
                        If oBitmap.GBPalette.Colors(i + (PalID * 4)).Red = pixelRGB.Red And oBitmap.GBPalette.Colors(i + (PalID * 4)).Green = pixelRGB.Green And oBitmap.GBPalette.Colors(i + (PalID * 4)).Blue = pixelRGB.Blue Then
                            bestColor = i - 1
                            Exit For
                        End If
                    Next i

                    oBitmap.PixelData((X * 8) + PixelX, (Y * 8) + PixelY) = bestColor

                Next PixelX
            Next PixelY

            Progress.value = Progress.value + 1
        Next X
    Next Y

End Sub

Private Sub chkBatch_Click()

    If chkBatch.value = vbChecked Then
        
        fraBatch.Enabled = True
        lblLabel(3).Enabled = True
        txtInputDir.Enabled = True
        cmdBrowseInput.Enabled = True
        FileList.Enabled = True
        chkOpenFiles.Enabled = True
        chkOpenFiles.value = vbUnchecked
        chkPackToGB.Enabled = True
    
    Else
    
        fraBatch.Enabled = False
        lblLabel(3).Enabled = False
        txtInputDir.Enabled = False
        cmdBrowseInput.Enabled = False
        FileList.Enabled = False
        chkOpenFiles.Enabled = False
        chkOpenFiles.value = vbChecked
        chkPackToGB.Enabled = False
    End If
    

End Sub

Private Sub chkBitmapOnly_Click()

    If chkBitmapOnly.value = vbChecked Then
        
        chkBatch.value = vbChecked
        chkOpenFiles.value = vbUnchecked
        chkPackToGB.value = vbChecked
        
        chkPalette.Enabled = False
        chkVRAM.Enabled = False
        chkPattern.Enabled = False
        txtPalPrefix.Enabled = False
        txtPatPrefix.Enabled = False
        txtVrmPrefix.Enabled = False
        optVRAMType(0).Enabled = False
        optVRAMType(1).Enabled = False
        cmdBrowseVRAM.Enabled = False
        optBGType(0).Enabled = False
        optBGType(1).Enabled = False
        chkBatch.Enabled = False
        chkOpenFiles.Enabled = False
        chkPackToGB.Enabled = False
        
    Else
        
        chkBatch.value = vbUnchecked
        chkOpenFiles.value = vbChecked
        chkPackToGB.value = vbUnchecked
    
        chkPalette.Enabled = True
        chkVRAM.Enabled = True
        chkPattern.Enabled = True
        txtPalPrefix.Enabled = True
        txtPatPrefix.Enabled = True
        txtVrmPrefix.Enabled = True
        optVRAMType(0).Enabled = True
        optVRAMType(1).Enabled = True
        cmdBrowseVRAM.Enabled = True
        optBGType(0).Enabled = True
        optBGType(1).Enabled = True
        chkBatch.Enabled = True
        chkOpenFiles.Enabled = True
        chkPackToGB.Enabled = True
    
    End If

End Sub

Private Sub chkPalette_Click()

    If chkPalette.value = vbChecked Then
        chkVRAM.Enabled = True
        txtPalPrefix.Enabled = True
    Else
        chkVRAM.Enabled = False
        txtPalPrefix.Enabled = False
        chkVRAM.value = vbUnchecked
    End If

End Sub

Private Sub chkPattern_Click()

    If chkPattern.value = vbChecked Then
        optBGType(0).Enabled = True
        optBGType(1).Enabled = True
        txtPatPrefix.Enabled = True
    Else
        optBGType(0).Enabled = False
        optBGType(1).Enabled = False
        txtPatPrefix.Enabled = False
    End If

End Sub

Private Sub chkVRAM_Click()

    If chkVRAM.value = vbChecked Then
        chkPattern.Enabled = True
        
        optVRAMType(0).Enabled = True
        optVRAMType(1).Enabled = True
        cmdBrowseVRAM.Enabled = True
        txtVRAMFile.Enabled = True
        txtVrmPrefix.Enabled = True
    
    Else
        chkPattern.Enabled = False
        chkPattern.value = vbUnchecked
        
        optVRAMType(0).value = True
        
        optVRAMType(0).Enabled = False
        optVRAMType(1).Enabled = False
        cmdBrowseVRAM.Enabled = False
        txtVRAMFile.Enabled = False
        txtVrmPrefix.Enabled = False
        
    End If

End Sub



Private Sub cmdBrowse_Click()

    Dim str As String
    
    str = UCase$(GetPathDialog & "\")

    If str <> "\" Then
        txtOutputDir.Text = str
    End If

End Sub

Private Sub cmdBrowseInput_Click()

    txtInputDir.Text = UCase$(GetPathDialog & "\")
    
End Sub


Private Sub cmdBrowseVRAM_Click()

    With mdiMain.Dialog
        .DialogTitle = "Specify GB VRAM File"
        .Filename = ""
        .Filter = "GB VRAMs (*.vrm)|*.vrm"
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        
        optVRAMType(0).value = True
                
        gsCurPath = .Filename
        
        txtVRAMFile.Text = .Filename
        
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


Private Sub cmdCreateFiles_Click()

'***************************************************************************
'   Load a GB Bitmap from a windows .bmp file
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    Dim str As String
    Dim flag As Boolean
    Dim d As String
    Dim sFilename As String
    
    If chkBatch.value = vbChecked Then
        If txtInputDir.Text = "" Then
            MsgBox "You must enter an input folder!", vbCritical, "Input Error"
            Exit Sub
        End If
    End If
    
    If optVRAMType(0).value = True Then
        If txtVRAMFile.Text = "" Then
            MsgBox "You must specify a VRAM File!", vbCritical, "Input Error"
            Exit Sub
        End If
    End If
    
    If FileList.Path <> "" Then
        str = FileList.Path
        If Mid$(str, Len(str), 1) <> "\" Then
            str = str & "\"
        End If
    End If

    If FileList.List(0) <> "" Then
        d = FileList.List(0)
    Else
        d = "      "
    End If
        
    Do Until FileList.ListIndex = FileList.ListCount Or d = ""
    
        If UCase$(Mid$(d, Len(d) - 3)) = ".BMP" Then
    
            If chkBatch.value = vbChecked Then
                sFilename = d
                FileList.ListIndex = FileList.ListIndex + 1
            Else
                GoTo Dialog
            End If
        Else
Dialog:
            With mdiMain.Dialog
            'Get filename
                .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
                .DialogTitle = "Get Windows Bitmap"
                .Filename = ""
                .Filter = "Windows Bitmap (*.bmp)|*.bmp"
                .ShowOpen
                If .Filename = "" Then
                    Exit Sub
                End If
                gsCurPath = .Filename
                sFilename = .Filename
                str = ""
                flag = True
            End With
        End If
            
        d = FileList.List(FileList.ListIndex + 1)
         
        mCreateResources str & sFilename, txtVRAMFile.Text
        
        If flag Then Exit Do

    Loop

    Progress.value = 0
    Status.Panels(1).Text = ""
    Screen.MousePointer = 0
    
    'MsgBox "All files created successfully!", vbInformation, "Success"
    
    gbBatchGoing = False
    
    Unload Me

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmAdvBitmap:cmdCreateFiles_Click Error"
End Sub

Private Sub FileList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    FileList.ListIndex = -1

End Sub


Private Sub Form_Load()

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
    Me.Show
    
    chkPalette.value = vbChecked
    chkVRAM.value = vbChecked
    chkPattern.value = vbChecked
    
End Sub


Private Sub txtInputDir_Change()
    
    FileList.Path = txtInputDir.Text
    FileList.Pattern = "*.bmp"

End Sub

Private Sub txtInputDir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub


Private Sub txtOutputDir_KeyPress(KeyAscii As Integer)

    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub




Private Sub txtPalPrefix_KeyPress(KeyAscii As Integer)

    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
    
End Sub


Private Sub txtPatPrefix_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub



Private Sub txtVrmPrefix_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub


