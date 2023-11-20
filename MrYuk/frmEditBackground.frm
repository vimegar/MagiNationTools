VERSION 5.00
Begin VB.Form frmEditBackground 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmEditBackground.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   Visible         =   0   'False
   Begin VB.CheckBox chkMode 
      Caption         =   "Edit Ti&les"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame fraSource 
      Caption         =   "Resour&ce Files"
      Height          =   2415
      Left            =   4200
      TabIndex        =   18
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmdBrowsePalette 
         Appearance      =   0  'Flat
         Caption         =   "Br&owse..."
         Height          =   360
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1920
         Width           =   900
      End
      Begin VB.CommandButton cmdEditPalette 
         Appearance      =   0  'Flat
         Caption         =   "Edi&t..."
         Height          =   360
         Left            =   1080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   900
      End
      Begin VB.CommandButton cmdBrowseVRAM 
         Appearance      =   0  'Flat
         Caption         =   "B&rowse..."
         Height          =   360
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmdEditVRAM 
         Appearance      =   0  'Flat
         Caption         =   "&Edit..."
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   900
      End
      Begin VB.Label txtPaletteSource 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1845
      End
      Begin VB.Label txtVRAMSource 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label lblPaletteSource 
         AutoSize        =   -1  'True
         Caption         =   "&Palette:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblVRAMSource 
         AutoSize        =   -1  'True
         Caption         =   "&VRAM:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As..."
      Height          =   360
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Frame fraTileOptions 
      Height          =   2055
      Left            =   4200
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox txtPriority 
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdYFlip 
         Caption         =   "&Y Flip"
         Height          =   360
         Left            =   1080
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1560
         Width           =   900
      End
      Begin VB.CommandButton cmdXFlip 
         Caption         =   "&X Flip"
         Height          =   360
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1560
         Width           =   900
      End
      Begin VB.TextBox txtPalID 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Priority:"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   26
         Top             =   960
         Width           =   510
      End
      Begin VB.Label txtBank 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label txtTileID 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "A&ddress:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Palette &ID:"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Bank:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1080
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   0
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   3840
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   255
      LargeChange     =   8
      Left            =   0
      Max             =   255
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3855
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   3855
      LargeChange     =   8
      Left            =   3840
      Max             =   255
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame fraEmpty 
      Height          =   2055
      Left            =   4200
      TabIndex        =   21
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Frame fraData 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2400
      TabIndex        =   27
      Top             =   4200
      Width           =   1770
      Begin VB.Label lblI 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1230
         TabIndex        =   37
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "I:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   8
         Left            =   1125
         TabIndex        =   36
         Top             =   15
         Width           =   75
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "W:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   7
         Left            =   585
         TabIndex        =   35
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   6
         Left            =   600
         TabIndex        =   34
         Top             =   195
         Width           =   105
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   5
         Left            =   60
         TabIndex        =   33
         Top             =   -15
         Width           =   135
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   4
         Left            =   60
         TabIndex        =   32
         Top             =   180
         Width           =   135
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblY 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   30
         Top             =   180
         Width           =   360
      End
      Begin VB.Label lblW 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   735
         TabIndex        =   29
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblH 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   735
         TabIndex        =   28
         Top             =   180
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmEditBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'   VB Interface Setup
'***************************************************************************
    
    Option Explicit
    Implements intResourceClient

'***************************************************************************
'   Form dimensions
'***************************************************************************

    Private Const DEF_WIDTH = 6570 'In Twips
    Private Const DEF_HEIGHT = 5055 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mViewport As New clsViewport
    Private mSelection As New clsSelection
    
    Private mbOpening As Boolean
    Private mbChanged As Boolean
    Private mbCached As Boolean
    Private msFilename As String
    Private mnSelectedPalette As Integer
    Private mnGridX As Integer
    Private mnGridY As Integer
    Private mbNeedApply As Boolean
    
    Private mbClipFlipX As Boolean
    Private mbClipFlipY As Boolean
    Private mnClipPal As Integer
    Private mSelBuffer As New clsOffscreen
    
'***************************************************************************
'   Resource object
'***************************************************************************

    Private mGBBackground As New clsGBBackground
Public Property Let bOpening(bNewValue As Boolean)

    mbOpening = bNewValue

End Property

Public Property Get GBBackground() As clsGBBackground

    Set GBBackground = mGBBackground

End Property

Public Property Set GBBackground(oNewValue As clsGBBackground)

    Set mGBBackground = oNewValue

End Property

Private Sub mApply()

'    On Error GoTo HandleErrors
'
'    If mbNeedApply = False Then
'        Exit Sub
'    End If
'
'    Dim ret As Long
'    Dim indexX As Integer
'    Dim indexY As Integer
'
'    ret = gSelection.GetFirstElement
'
'    Do Until ret < 0
'
'        indexX = gSelection.Left + gSelection.CursorX + 1
'        indexY = gSelection.Top + gSelection.CursorY + 1
'
'        With mGBBackground
'            .VRAMEntryBank(indexX, indexY) = val(txtBank.Caption)
'            .PaletteID(indexX, indexY) = val(txtPalID.Text)
'            .BitmapFragmentIndex(indexX, indexY) = mGBBackground.GBVRAM.GetBitFragIDFromVRAMAddr(.VRAMEntryAddress(indexX, indexY))
'        End With
'        ret = gSelection.GetNextElement
'    Loop
'
'    intResourceClient_Update

'Exit Sub

'HandleErrors:
'    MsgBox Err.Description, vbCritical, "frmEditBackground:mApply Error"
End Sub

Public Property Get nGridX() As Integer

    nGridX = mnGridX

End Property

Public Property Get nGridY() As Integer

    nGridY = mnGridY

End Property

Public Property Get nSelectedPalette() As Integer

    nSelectedPalette = mnSelectedPalette

End Property

Public Property Let nSelectedPalette(nNewValue As Integer)

    mnSelectedPalette = nNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Private Sub mDrawGridLinesOnBank()

'***************************************************************************
'   Create grid on display window
'***************************************************************************

    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim xCount As Integer
    Dim yCount As Integer
    
    For Y = 1 To mGBBackground.nHeight * 8 Step 16
        
        xCount = 0
        
        For X = 1 To mGBBackground.nWidth * 8 Step 16
        
        'Draw the large rectangles with bright color
            mGBBackground.MapOffscreen.RECT X, Y, X + 16, Y + 16, RGB(196, 141, 180)
            mGBBackground.MapOffscreen.RECT X + 1, Y + 1, X + 15, Y + 15, RGB(113, 64, 99)
        
        'Draw the smaller, lighter rectangles
            For i = 1 To 15
                mGBBackground.MapOffscreen.SetPixel i + X, 8 + 16 * yCount, RGB(147, 106, 135)
                mGBBackground.MapOffscreen.SetPixel 8 + 16 * xCount, i + Y, RGB(147, 106, 135)
            Next i
            
            xCount = xCount + 1
            
        Next X
        
        yCount = yCount + 1
        
    Next Y

End Sub



Private Sub chkMode_Click()

    Screen.MousePointer = vbHourglass

    Dim bTempFlag As Boolean
    bTempFlag = mbChanged

    If chkMode.value = vbChecked Then
        
        fraTileOptions.Visible = True
        fraEmpty.Visible = False
        mUpdateMore
        
        mnGridX = 8
        mnGridY = 8
    
        mSelection.Left = gSelection.Left * 2
        mSelection.Top = gSelection.Top * 2
        mSelection.Right = gSelection.Right * 2
        mSelection.Bottom = gSelection.Bottom * 2
    
    Else

        fraTileOptions.Visible = False
        fraEmpty.Visible = True
    
        mbNeedApply = False
        
        mnGridX = 16
        mnGridY = 16
    
        mbNeedApply = True
    
        mSelection.Left = gSelection.Left / 2
        mSelection.Top = gSelection.Top / 2
        mSelection.Right = gSelection.Right / 2
        mSelection.Bottom = gSelection.Bottom / 2
    
    End If

    mSelection.AreaWidth = (mGBBackground.nWidth * 8) \ mnGridX
    mSelection.AreaHeight = (mGBBackground.nHeight * 8) \ mnGridY
    mSelection.CellWidth = mnGridX
    mSelection.CellHeight = mnGridY
    Set gSelection = mSelection.FixRect
    Set gSelection.SrcForm = Me
        
    mbChanged = bTempFlag
    
    intResourceClient_Update

    Screen.MousePointer = 0

End Sub



Private Sub cmdBrowsePalette_Click()
    
'***************************************************************************
'   Get the filename for the palette
'***************************************************************************

    On Error GoTo HandleErrors
    
    With mdiMain.Dialog
        .InitDir = mGBBackground.intResource_ParentPath & "\Palettes"
        .DialogTitle = "Load Palette"
        .Filename = ""
        .Filter = "GB Palettes (*.pal)|*.pal"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        
        Dim sDir As String
        Dim sFilename As String
        sDir = "Palettes"
        sFilename = .Filename
        Set mGBBackground.GBPalette = gResourceCache.GetResourceFromFile(sFilename, mGBBackground)
    
        gbOpeningChild = True
        mGBBackground.sPaletteFile = sFilename
        gbOpeningChild = False
    
    End With

    intResourceClient_Update

Exit Sub

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sFilename & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = mGBBackground.intResource_ParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                Exit Sub
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, mGBBackground.intResource_ParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, mGBBackground.intResource_ParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
                        
            Resume
            
        End With
    
    Else
        MsgBox Err.Description, vbCritical, "frmEditBackground:cmdBrowsePalette_Click Error"
    End If
End Sub

Private Sub cmdBrowseVRAM_Click()

'***************************************************************************
'   Get the filename for the VRAM
'***************************************************************************
 
    On Error GoTo HandleErrors
    
    With mdiMain.Dialog
        .InitDir = mGBBackground.intResource_ParentPath & "\VRAMs"
        .DialogTitle = "Load VRAM"
        .Filename = ""
        .Filter = "GB VRAMs (*.vrm)|*.vrm"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        
        Dim sDir As String
        Dim sFilename As String
        
        sDir = "VRAMs"
        sFilename = .Filename
        
        gbOpeningChild = True
        Set mGBBackground.GBVRAM = gResourceCache.GetResourceFromFile(sFilename, mGBBackground)
        gbOpeningChild = False
    
        mGBBackground.sVRAMFile = sFilename
    
    End With
    
    intResourceClient_Update

Exit Sub

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sFilename & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = mGBBackground.intResource_ParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                Exit Sub
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, mGBBackground.intResource_ParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, mGBBackground.intResource_ParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
                        
            Resume
            
        End With
    
    Else

        MsgBox Err.Description, vbCritical, "frmEditBackground:cmdBrowseVRAM_Click Error"
    End If
End Sub

Private Sub cmdEditPalette_Click()

'***************************************************************************
'   Open the palette file in an editor
'***************************************************************************

    On Error GoTo HandleErrors

    Screen.MousePointer = vbHourglass

    If mGBBackground.GBPalette Is Nothing Then
        MsgBox "No Palette file specified!", vbCritical, "Error"
        Exit Sub
    End If
    
    Dim frm As New frmEditPalette
    gResourceCache.AddResourceToCache mGBBackground.sPaletteFile, mGBBackground.GBPalette, frm
    Set frm.GBPalette = mGBBackground.GBPalette
    frm.sFilename = mGBBackground.intResource_ParentPath & "\Palettes\" & GetTruncFilename(mGBBackground.sPaletteFile)
    frm.Show

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdEditPalette_Click Error"
    Screen.MousePointer = 0
End Sub

Private Sub cmdEditVRAM_Click()

'***************************************************************************
'   Open the VRAM file in an editor
'***************************************************************************

    On Error GoTo HandleErrors

    Screen.MousePointer = vbHourglass

    If mGBBackground.GBVRAM Is Nothing Then
        MsgBox "No VRAM file specified!", vbCritical, "Error"
        Exit Sub
    End If
    
    Dim frm As New frmEditVRAM
    gResourceCache.AddResourceToCache mGBBackground.sVRAMFile, mGBBackground.GBVRAM, frm
    Set frm.GBVRAM = mGBBackground.GBVRAM
    frm.sFilename = mGBBackground.intResource_ParentPath & "\VRAMs\" & GetTruncFilename(mGBBackground.sVRAMFile)
    frm.Show

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdEditVRAM_Click Error"
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()

'***************************************************************************
'   Save the current pattern into a .pat file
'***************************************************************************

    On Error GoTo HandleErrors

    If mGBBackground.tBackgroundType = GB_PATTERNBG Then
        PackFile mGBBackground.intResource_ParentPath & "\Patterns\" & GetTruncFilename(msFilename), mGBBackground
    Else
        PackFile mGBBackground.intResource_ParentPath & "\Backgrounds\" & GetTruncFilename(msFilename), mGBBackground
    End If
    
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdSave_Click Error"
End Sub

Private Sub cmdSaveAs_Click()

'***************************************************************************
'   Save the current pattern under a new filename
'***************************************************************************

    On Error GoTo HandleErrors

    Screen.MousePointer = vbHourglass

    With mdiMain.Dialog
        
    'Get filename
        If mGBBackground.tBackgroundType = GB_RAWBG Then
            .InitDir = gsCurPath
            .DialogTitle = "Save GB Background"
            .Filename = ""
            .Filter = "GB Backgrounds (*.bg)|*.bg"
        Else
            .InitDir = gsCurPath
            .DialogTitle = "Save GB Pattern"
            .Filename = ""
            .Filter = "GB Patterns (*.pat)|*.pat"
        End If
        
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBBackground
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBBackground.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBBackground, Me
        
    'Set flag used for saving
        
        intResourceClient_Update
        
        mbChanged = False

    End With

rExit:
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdSaveAs_Click Error"
    GoTo rExit
End Sub



Private Sub cmdXFlip_Click()

    On Error GoTo HandleErrors

    Dim srcX As Integer
    Dim srcY As Integer
    Dim destX As Integer
    Dim tempX As Integer
    Dim tempY As Integer
    Dim tempData As New clsGBBackground
    
    tempData.nWidth = mGBBackground.nWidth
    tempData.nHeight = mGBBackground.nHeight
    
    For tempY = 1 To mGBBackground.nHeight
        For tempX = 1 To mGBBackground.nWidth
            tempData.VRAMEntryAddress(tempX, tempY) = mGBBackground.VRAMEntryAddress(tempX, tempY)
            tempData.VRAMEntryBank(tempX, tempY) = mGBBackground.VRAMEntryBank(tempX, tempY)
            tempData.BitmapFragmentIndex(tempX, tempY) = mGBBackground.BitmapFragmentIndex(tempX, tempY)
            tempData.PaletteID(tempX, tempY) = mGBBackground.PaletteID(tempX, tempY)
            tempData.Priority(tempX, tempY) = mGBBackground.Priority(tempX, tempY)
            tempData.YFlip(tempX, tempY) = mGBBackground.YFlip(tempX, tempY)
        Next tempX
    Next tempY
    
    For srcY = gSelection.Bottom To gSelection.Top + 1 Step -1
        destX = gSelection.Left
        For srcX = gSelection.Right To gSelection.Left + 1 Step -1
            destX = destX + 1
        
            With mGBBackground
                .VRAMEntryAddress(destX, srcY) = tempData.VRAMEntryAddress(srcX, srcY)
                .VRAMEntryBank(destX, srcY) = tempData.VRAMEntryBank(srcX, srcY)
                .BitmapFragmentIndex(destX, srcY) = tempData.BitmapFragmentIndex(srcX, srcY)
                .PaletteID(destX, srcY) = tempData.PaletteID(srcX, srcY)
                .Priority(destX, srcY) = tempData.Priority(srcX, srcY)
                .XFlip(destX, srcY) = -(.XFlip(destX, srcY) = 0)
                .YFlip(destX, srcY) = tempData.YFlip(srcX, srcY)
            End With
        
        Next srcX
    Next srcY

    mbChanged = True
    mbClipFlipX = True
    intResourceClient_Update
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdXFlip_Click Error"
End Sub

Private Sub cmdYFlip_Click()

    On Error GoTo HandleErrors

    Dim srcX As Integer
    Dim srcY As Integer
    Dim destY As Integer
    Dim tempX As Integer
    Dim tempY As Integer
    Dim tempData As New clsGBBackground
    
    tempData.nWidth = mGBBackground.nWidth
    tempData.nHeight = mGBBackground.nHeight
    
    For tempY = 1 To mGBBackground.nHeight
        For tempX = 1 To mGBBackground.nWidth
            tempData.VRAMEntryAddress(tempX, tempY) = mGBBackground.VRAMEntryAddress(tempX, tempY)
            tempData.VRAMEntryBank(tempX, tempY) = mGBBackground.VRAMEntryBank(tempX, tempY)
            tempData.BitmapFragmentIndex(tempX, tempY) = mGBBackground.BitmapFragmentIndex(tempX, tempY)
            tempData.PaletteID(tempX, tempY) = mGBBackground.PaletteID(tempX, tempY)
            tempData.Priority(tempX, tempY) = mGBBackground.Priority(tempX, tempY)
            tempData.XFlip(tempX, tempY) = mGBBackground.XFlip(tempX, tempY)
        Next tempX
    Next tempY
    
    For srcX = gSelection.Right To gSelection.Left + 1 Step -1
        destY = gSelection.Top
        For srcY = gSelection.Bottom To gSelection.Top + 1 Step -1
            destY = destY + 1
        
            With mGBBackground
                .VRAMEntryAddress(srcX, destY) = tempData.VRAMEntryAddress(srcX, srcY)
                .VRAMEntryBank(srcX, destY) = tempData.VRAMEntryBank(srcX, srcY)
                .BitmapFragmentIndex(srcX, destY) = tempData.BitmapFragmentIndex(srcX, srcY)
                .PaletteID(srcX, destY) = tempData.PaletteID(srcX, srcY)
                .Priority(srcX, destY) = tempData.Priority(srcX, srcY)
                .XFlip(srcX, destY) = tempData.XFlip(srcX, srcY)
                .YFlip(srcX, destY) = -(.YFlip(srcX, destY) = 0)
            End With
        
        Next srcY
    Next srcX

    mbChanged = True
    mbClipFlipY = True
    intResourceClient_Update
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:cmdYFlip_Click Error"
End Sub






Private Sub Form_Load()
    
'***************************************************************************
'   Load the pattern editor
'***************************************************************************

    On Error GoTo HandleErrors
    
'Set form dimensions
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
        
    mnGridX = 8
    mnGridY = 8
        
    If Not mbOpening Then
    'Display bg properties properties
        frmBGProperties.nWidth = 32
        frmBGProperties.nHeight = 32
        frmBGProperties.Show vbModal
    
        If frmBGProperties.bCancel Then
            gbClosingApp = True
            mbChanged = False
            Unload frmBGProperties
            Unload Me
            gbClosingApp = False
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        Screen.MousePointer = vbHourglass
        
        mGBBackground.nWidth = frmBGProperties.nWidth
        mGBBackground.nHeight = frmBGProperties.nHeight
    
        gbClosingApp = True
        Unload frmBGProperties
        gbClosingApp = False
        
        Screen.MousePointer = 0
        
    End If

    If mGBBackground.tBackgroundType = GB_PATTERNBG Then

        'chkMode.Enabled = True
        mnGridX = 16
        mnGridY = 16

    End If
    
    mSelection.Left = 0
    mSelection.Top = 0
    mSelection.Right = mSelection.Left + 1
    mSelection.Bottom = mSelection.Top + 1
    mSelection.AreaWidth = (mGBBackground.nWidth * 8) \ mnGridX
    mSelection.AreaHeight = (mGBBackground.nHeight * 8) \ mnGridY
    mSelection.CellWidth = mnGridX
    mSelection.CellHeight = mnGridY
            
    Set gSelection = mSelection.FixRect
    Set gSelection.SrcForm = Me
    
'Set up visual objects
    
    mGBBackground.Offscreen.Create mGBBackground.nWidth * 8, mGBBackground.nHeight * 8
    mGBBackground.MapOffscreen.Create mGBBackground.nWidth * 8, mGBBackground.nHeight * 8
    
    mSelBuffer.Create mGBBackground.nWidth * 8, mGBBackground.nHeight * 8
    
    With mViewport
        .Zoom = 1
        
        .ViewportWidth = 256
        .ViewportHeight = 256
        
        .SourceWidth = mGBBackground.nWidth * 8
        .SourceHeight = mGBBackground.nHeight * 8
        
        .DisplayWidth = 256
        .DisplayHeight = 256
    
        hsbScroll.Max = .hScrollMax
        vsbScroll.Max = .vScrollMax
    End With
    
'Set up bitmap fragments
    mGBBackground.GBVRAM.EnumBitmapFragments
    
    lblW.Caption = Format$(CStr(mGBBackground.nWidth), "000")
    lblH.Caption = Format$(CStr(mGBBackground.nHeight), "000")
    
'Update visual
    intResourceClient_Update

'Organize forms on screen
    CleanUpForms Me

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:Form_Load Error"
End Sub

Private Sub mUpdateMore()

    On Error GoTo HandleErrors

    If gSelection.SrcForm Is Nothing Then
        Exit Sub
    End If
    
    If Not TypeOf gSelection.SrcForm Is frmEditBackground Then
        Exit Sub
    End If

'Refresh the data to be displayed
    If gSelection.Left < 0 Or gSelection.Top < 0 Then
        Exit Sub
    End If
    
    mbNeedApply = False
    
    With mGBBackground
        If gSelection.SelectionWidth > 1 Or gSelection.SelectionHeight > 1 Or gSelection.Left > mGBBackground.nWidth Or gSelection.Top > mGBBackground.nHeight Then
        Else
            txtTileID.Caption = Hex(CStr(.VRAMEntryAddress(gSelection.Left + 1, gSelection.Top + 1)))
            txtBank.Caption = CStr(.VRAMEntryBank(gSelection.Left + 1, gSelection.Top + 1))
            txtPalID.Text = CStr(.PaletteID(gSelection.Left + 1, gSelection.Top + 1))
            txtPriority.Text = CStr(.Priority(gSelection.Left + 1, gSelection.Top + 1))
        End If
    End With

    mbNeedApply = True

Exit Sub

HandleErrors:
    If Err.Description = "Subscript out of range" Then
        MsgBox "(" & Err.Description & ")" & vbCrLf & "The area you have clicked on is most likely out of the bounds of the current BG!", vbCritical, "frmEditBackground:mUpdateMore Error"
    Else
        MsgBox Err.Description, vbCritical, "frmEditBackground:mUpdateMore Error"
    End If
End Sub

Public Property Get bChanged() As Boolean

    bChanged = mbChanged

End Property

Public Property Let bChanged(bNewValue As Boolean)

    mbChanged = bNewValue

End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Select Case gnTool
    
        Case GB_POINTER
        
            Me.MousePointer = vbArrow
        
        Case GB_MARQUEE
        
            Me.MousePointer = vbArrow
        
        Case GB_ZOOM
    
            Me.MousePointer = vbArrow
                
        Case GB_BRUSH
            
            Me.MousePointer = vbArrow
            
        Case GB_SETTER
        
            Me.MousePointer = vbArrow
        
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'***************************************************************************
'   Confirm file saving just before the form is closed
'***************************************************************************

    If mbChanged = True Then
        
    'Prompt the user
        Dim ret As Integer
        ret = MsgBox("Do you want to save " & GetTruncFilename(msFilename) & " before closing?", vbQuestion Or vbYesNoCancel, "Confirmation")
        
    'Save or cancel: whichever is appropriate
        If ret = vbYes Then
            cmdSave_Click
        ElseIf ret = vbCancel Then
            Cancel = True
        End If
        
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
'***************************************************************************
'   Release memory when form closes
'***************************************************************************

    On Error GoTo HandleErrors
    
    gResourceCache.ReleaseClient Me
    
    If Not mSelBuffer Is Nothing Then
        mSelBuffer.Delete
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:Form_Unload Error"
End Sub

Private Sub hsbScroll_Change()

    On Error GoTo HandleErrors

    If mGBBackground.tBackgroundType = GB_RAWBG Then
        picDisplay.Cls
    End If

    mViewport.ViewportX = hsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBBackground.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:hsbScroll_Change Error"

End Sub

Private Sub hsbScroll_Scroll()

    hsbScroll_Change
    
End Sub


Public Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

'***************************************************************************
'   Update the visual display
'***************************************************************************

    If mGBBackground.tBackgroundType = GB_RAWBG Then
        picDisplay.Cls
    End If

'Draw grid on the display
    mDrawGridLinesOnBank
    
'Draw bitmap fragments
    If Not mGBBackground.GBVRAM Is Nothing Then
        
        Dim nx As Integer
        Dim ny As Integer
    
        For ny = 1 To mGBBackground.nHeight
            For nx = 1 To mGBBackground.nWidth
                If mGBBackground.GBVRAM.GetBitFragIDFromVRAMAddr(mGBBackground.VRAMEntryAddress(nx, ny)) <= 384 Then
                    With mGBBackground.GBVRAM.BitmapFragments(mGBBackground.GBVRAM.GetBitFragIDFromVRAMAddr(mGBBackground.VRAMEntryAddress(nx, ny)), mGBBackground.VRAMEntryBank(nx, ny))
                        'DoEvents
                        If Not .GBBitmap Is Nothing Then
                            .GBBitmap.BlitWithPal mGBBackground.MapOffscreen.hdc, (nx - 1) * 8, (ny - 1) * 8, 8, 8, CLng(.X), CLng(.Y), mGBBackground.GBPalette, mGBBackground.PaletteID(nx, ny), mGBBackground.XFlip(nx, ny), mGBBackground.YFlip(nx, ny)
                        End If
                    End With
                End If
            Next nx
        Next ny
        
    End If
    
    mGBBackground.MapOffscreen.BlitRect mGBBackground.Offscreen.hdc, 0, 0, mGBBackground.Offscreen.width, mGBBackground.Offscreen.height, 0, 0
    mGBBackground.Offscreen.Blit mSelBuffer.hdc, 0, 0
    
'If using the marquee tool, show the selected area inversed
    If gSelection.SrcForm Is Me Then
        mGBBackground.Offscreen.BlitRaster mSelBuffer.hdc, gSelection.Left * gSelection.CellWidth, gSelection.Top * gSelection.CellHeight, gSelection.SelectionWidth * gSelection.CellWidth, gSelection.SelectionHeight * gSelection.CellHeight, 0, 0, vbDstInvert
    End If
    
'Update the screen
    mViewport.Draw picDisplay.hdc, mSelBuffer
    
    Me.Caption = GetTruncFilename(msFilename)
    txtPaletteSource.Caption = GetTruncFilename(mGBBackground.sPaletteFile)
    txtVRAMSource.Caption = GetTruncFilename(mGBBackground.sVRAMFile)
    
    picDisplay.Refresh
    
    If mbOpening Then
        mbOpening = False
    Else
        If tType = GB_ACTIVEEDITOR Then
            If Not mGBBackground Is Nothing Then
                mGBBackground.UpdateClients Me
                mGBBackground.UpdateClients Me
            End If
        End If
    
        If mGBBackground.tBackgroundType = GB_PATTERNBG Or (mGBBackground.tBackgroundType = GB_RAWBG And chkMode.value = vbChecked) Then
            mUpdateMore
        End If
    End If

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:intResourceClient_Update Error"
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Dim i As Integer
    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), mnGridX, mnGridY)
    
    Select Case gnTool
        
        Case GB_POINTER
lPointer:
            'If mGBBackground.tBackgroundType = GB_RAWBG And chkMode.value = vbUnchecked Then
            '    picDisplay.MousePointer = vbArrow
            '    Exit Sub
            'End If
            
        'Create a selection where the user clicked
            
            mSelection.Left = click.X
            mSelection.Top = click.Y
            mSelection.Right = mSelection.Left + 1
            mSelection.Bottom = mSelection.Top + 1
            mSelection.AreaWidth = (mGBBackground.nWidth * 8) \ mnGridX
            mSelection.AreaHeight = (mGBBackground.nHeight * 8) \ mnGridY
            mSelection.CellWidth = mnGridX
            mSelection.CellHeight = mnGridY
            
            Dim rclick As New clsPoint
            Set rclick = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)
            gnReplaceTile = rclick.X + (rclick.Y * (mGBBackground.nWidth \ 2))
            
            Set gSelection = mSelection.FixRect
            Set gSelection.SrcForm = Me
            
            mPasteSel
            
            If chkMode.value = vbUnchecked Then
                SelectTool GB_BRUSH
            End If
                
        Case GB_MARQUEE
        
            'If mGBBackground.tBackgroundType = GB_RAWBG Then
            '    picDisplay.MousePointer = vbArrow
            '    Exit Sub
            'End If
        
        'Create a selection where the user clicked
            mSelection.Left = click.X
            mSelection.Top = click.Y
            mSelection.Right = mSelection.Left + 1
            mSelection.Bottom = mSelection.Top + 1
            mSelection.AreaWidth = (mGBBackground.nWidth * 8) \ mnGridX
            mSelection.AreaHeight = (mGBBackground.nHeight * 8) \ mnGridY
            mSelection.CellWidth = mnGridX
            mSelection.CellHeight = mnGridY
            
            Set rclick = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)
            gnReplaceTile = rclick.X + (rclick.Y * (mGBBackground.nWidth \ 2))
            
            Set gSelection = mSelection.FixRect
            Set gSelection.SrcForm = Me
            
            mPasteSel
        
        Case GB_BRUSH
            
            If chkMode.value <> vbChecked Then
                SelectTool GB_POINTER
                GoTo lPointer
                Exit Sub
            End If
            
            Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 8, 8)
            
            If Button = vbRightButton Then 'Erase
                With mGBBackground
                    .VRAMEntryAddress(click.X + 1, click.Y + 1) = 0
                    .BitmapFragmentIndex(click.X + 1, click.Y + 1) = 0
                    .Priority(click.X + 1, click.Y + 1) = 0
                    .PaletteID(click.X + 1, click.Y + 1) = 0
                    .VRAMEntryBank(click.X + 1, click.Y + 1) = 0
                    .XFlip(click.X + 1, click.Y + 1) = 0
                    .YFlip(click.X + 1, click.Y + 1) = 0
                End With
                intResourceClient_Update
                Exit Sub
            End If
            
            If Not (TypeOf gSelection.SrcForm Is frmEditVRAM Or TypeOf gSelection.SrcForm Is frmEditBackground) Then
                SelectTool GB_POINTER
                picDisplay_MouseDown Button, Shift, X, Y
                SelectTool GB_POINTER
                GoTo lPointer
                Exit Sub
            End If
            
            Dim ret As Integer
            Dim xIndex As Integer
            Dim yIndex As Integer
            Dim retX As Integer
            Dim retY As Integer
            Dim d As Integer
            Dim retW As Integer
            Dim ix As Integer
            Dim iy As Integer
                        
            If TypeOf gSelection.SrcForm Is frmEditBackground Then
    
            'Transfer data from source to pattern
                ret = gSelection.GetFirstElement
                
                Do Until ret < 0
                    xIndex = click.X + (gSelection.CursorX * (gSelection.SrcForm.nGridX \ 8))
                    yIndex = click.Y + (gSelection.CursorY * (gSelection.SrcForm.nGridY \ 8))
                    If xIndex < 0 Or xIndex > mGBBackground.nWidth - 1 Or yIndex < 0 Or yIndex > mGBBackground.nHeight - 1 Then
                    Else
                        For iy = 1 To (gSelection.SrcForm.nGridY \ 8)
                            For ix = 1 To (gSelection.SrcForm.nGridX \ 8)
                                retW = (mGBBackground.nWidth * 8) \ gSelection.SrcForm.nGridX
                                retX = ((((ret - 1) * (gSelection.SrcForm.nGridX \ 8)) + ix) - 1) Mod mGBBackground.nWidth
                                retX = retX + 1
                                If retX <= mGBBackground.nWidth \ (gSelection.SrcForm.nGridX \ 8) Then
                                    retY = ((((ret - 1) * (gSelection.SrcForm.nGridY \ 8))) \ retW) + iy
                                Else
                                    retY = ((((ret - 1) * (gSelection.SrcForm.nGridY \ 8))) \ retW) + iy - 1
                                End If
                                On Error Resume Next
                                
                                mGBBackground.VRAMEntryBank(xIndex + ix, yIndex + iy) = gBufferBackground.VRAMEntryBank(retX, retY)
                                mGBBackground.VRAMEntryAddress(xIndex + ix, yIndex + iy) = gBufferBackground.VRAMEntryAddress(retX, retY)
                                mGBBackground.BitmapFragmentIndex(xIndex + ix, yIndex + iy) = gBufferBackground.BitmapFragmentIndex(retX, retY)
                                mGBBackground.PaletteID(xIndex + ix, yIndex + iy) = gBufferBackground.PaletteID(retX, retY)
                                mGBBackground.Priority(xIndex + ix, yIndex + iy) = gBufferBackground.Priority(retX, retY)
                                mGBBackground.XFlip(xIndex + ix, yIndex + iy) = gBufferBackground.XFlip(retX, retY)
                                mGBBackground.YFlip(xIndex + ix, yIndex + iy) = gBufferBackground.YFlip(retX, retY)

                                On Error GoTo HandleErrors
                            Next ix
                        Next iy
                    End If
                    ret = gSelection.GetNextElement
                Loop
                
            ElseIf TypeOf gSelection.SrcForm Is frmEditVRAM Then
            'Transfer data from source to pattern
                Set mGBBackground.GBVRAM = gSelection.SrcForm.GBVRAM
                
                ret = gSelection.GetFirstElement
                
                Do Until ret < 0
                    xIndex = click.X + gSelection.CursorX + 1
                    yIndex = click.Y + gSelection.CursorY + 1
                    If xIndex < 1 Or xIndex > mGBBackground.nWidth Or yIndex < 1 Or yIndex > mGBBackground.nHeight Then
                    Else
                        mGBBackground.VRAMEntryBank(xIndex, yIndex) = gSelection.SrcForm.SelectedBank
                        mGBBackground.VRAMEntryAddress(xIndex, yIndex) = ((ret - 1) * 16) + 32768
                        mGBBackground.BitmapFragmentIndex(xIndex, yIndex) = mGBBackground.GBVRAM.GetBitFragIDFromVRAMAddr(mGBBackground.VRAMEntryAddress(xIndex, yIndex))
                        mGBBackground.PaletteID(xIndex, yIndex) = 0
                        mGBBackground.Priority(xIndex, yIndex) = 0
                        mGBBackground.XFlip(xIndex, yIndex) = 0
                        mGBBackground.YFlip(xIndex, yIndex) = 0
                        
                    End If
                    ret = gSelection.GetNextElement
                Loop
                
            'Update VRAM Source text box
                If mGBBackground.sVRAMFile = "" Then
                    If gSelection.SrcForm.sFilename <> "" Then
                        mGBBackground.sVRAMFile = gSelection.SrcForm.sFilename
                    End If
                End If
            End If
        
        'Set up bitmap fragments
            mGBBackground.GBVRAM.EnumBitmapFragments
            
            intResourceClient_Update
            
        'Set variable used for saving
            mbChanged = True
            
        Case GB_ZOOM
            
            Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 1, 1)
            
            If Button = vbRightButton Then
                mViewport.ZoomOutToClick click.X, click.Y
            Else
                mViewport.ZoomInToClick click.X, click.Y
            End If
            
            hsbScroll.Max = mViewport.hScrollMax
            vsbScroll.Max = mViewport.vScrollMax
                
            If hsbScroll.Max >= mViewport.ViewportX Then
                hsbScroll.value = mViewport.ViewportX
            End If
            
            If vsbScroll.Max >= mViewport.ViewportY Then
                vsbScroll.value = mViewport.ViewportY
            End If
            
            intResourceClient_Update
        
        Case GB_SETTER
        
        'Set palette based on the currently selected palette
            If gPaletteSrcForm Is Nothing Then
                SelectTool GB_POINTER
                picDisplay_MouseDown Button, Shift, X, Y
                Exit Sub
            End If
            
            Set mGBBackground.GBPalette = gPaletteSrcForm.GBPalette
            
            Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 8, 8)

            If click.X < 0 Or click.Y < 0 Or click.X >= mGBBackground.nWidth Or click.Y >= mGBBackground.nHeight Then
                Exit Sub
            End If
            
            mGBBackground.PaletteID(click.X + 1, click.Y + 1) = gPaletteSrcForm.SelectedPalette
        
        'Update display
            intResourceClient_Update
        
        'Set variable used for saving
            mbChanged = True
        
        Case GB_BUCKET
        
            SelectTool GB_POINTER
            picDisplay_MouseDown Button, Shift, X, Y
            
        Case GB_REPLACE
        
            SelectTool GB_POINTER
            picDisplay_MouseDown Button, Shift, X, Y
        
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:picDisplay_MouseDown Error"
End Sub
Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

    If X >= picDisplay.width Or X < 0 Or Y >= picDisplay.height Or Y < 0 Then
        Exit Sub
    End If

    If Button = vbLeftButton And gnTool <> GB_ZOOM And gnTool <> GB_MARQUEE Then
        picDisplay_MouseDown Button, Shift, X, Y
    End If

    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), mnGridX, mnGridY)
    
    lblX.Caption = Format$(CStr(click.X), "000")
    lblY.Caption = Format$(CStr(click.Y), "000")
    
    Dim d As Integer
    If mnGridX = 16 Then
        d = 16
    ElseIf mnGridX = 8 Then
        d = 32
    End If
    
    lblI.Caption = Format$(CStr(click.X + (click.Y * d)), "0000")
    
'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
        
            picDisplay.MousePointer = vbArrow
        
        Case GB_MARQUEE
        
            'If mGBBackground.tBackgroundType = GB_RAWBG Then
            '    picDisplay.MousePointer = vbArrow
            '    Exit Sub
            'End If
            
            picDisplay.MousePointer = vbCrosshair
        
            If Button = vbLeftButton Then
                
                Dim dummyX As Integer
                Dim dummyY As Integer
                dummyX = 1
                dummyY = 1
                If mSelection.Right < mSelection.Left Then
                    dummyX = 0
                End If
                If mSelection.Bottom < mSelection.Top Then
                    dummyY = 0
                End If
                mSelection.Right = click.X + dummyX
                mSelection.Bottom = click.Y + dummyY
                
                Set gSelection = mSelection.FixRect
                Set gSelection.SrcForm = Me
                
                mPasteSel
            
            End If
        
        Case GB_BRUSH
            
            If TypeOf gSelection.SrcForm Is frmEditVRAM Or TypeOf gSelection.SrcForm Is frmEditBackground Then
                If chkMode.value <> vbChecked Then
                    picDisplay.MousePointer = vbArrow
                Else
                    picDisplay.MouseIcon = mdiMain.picDragAdd.Picture
                    picDisplay.MousePointer = vbCustom
                End If
            Else
                picDisplay.MousePointer = vbArrow
            End If
    
        Case GB_ZOOM
    
            picDisplay.MouseIcon = mdiMain.picZoom.Picture
            picDisplay.MousePointer = vbCustom
    
        Case GB_SETTER
        
            picDisplay.MousePointer = vbArrow
    
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:picDisplay_MouseMove Error"
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
            
            'If mGBBackground.tBackgroundType = GB_RAWBG Then
            '    picDisplay.MousePointer = vbArrow
            '    Exit Sub
            'End If
            
            Dim dx As Integer
            Dim dy As Integer
        
            gBufferBackground.nWidth = gSelection.SrcForm.GBBackground.nWidth
            gBufferBackground.nHeight = gSelection.SrcForm.GBBackground.nHeight
        
            For dy = 1 To gSelection.SrcForm.GBBackground.nHeight
                For dx = 1 To gSelection.SrcForm.GBBackground.nWidth
                
                    gBufferBackground.PaletteID(dx, dy) = gSelection.SrcForm.GBBackground.PaletteID(dx, dy)
                    gBufferBackground.Priority(dx, dy) = gSelection.SrcForm.GBBackground.Priority(dx, dy)
                    gBufferBackground.VRAMEntryAddress(dx, dy) = gSelection.SrcForm.GBBackground.VRAMEntryAddress(dx, dy)
                    gBufferBackground.VRAMEntryBank(dx, dy) = gSelection.SrcForm.GBBackground.VRAMEntryBank(dx, dy)
                    gBufferBackground.XFlip(dx, dy) = gSelection.SrcForm.GBBackground.XFlip(dx, dy)
                    gBufferBackground.YFlip(dx, dy) = gSelection.SrcForm.GBBackground.YFlip(dx, dy)
                    gBufferBackground.BitmapFragmentIndex(dx, dy) = gSelection.SrcForm.GBBackground.BitmapFragmentIndex(dx, dy)
                                
                Next dx
            Next dy
            
            mPasteSel
        
        Case GB_MARQUEE
        
            'If mGBBackground.tBackgroundType = GB_RAWBG Then
            '    picDisplay.MousePointer = vbArrow
            '    Exit Sub
            'End If
            
            gBufferBackground.nWidth = gSelection.SrcForm.GBBackground.nWidth
            gBufferBackground.nHeight = gSelection.SrcForm.GBBackground.nHeight
        
            For dy = 1 To gSelection.SrcForm.GBBackground.nHeight
                For dx = 1 To gSelection.SrcForm.GBBackground.nWidth
                
                    gBufferBackground.PaletteID(dx, dy) = gSelection.SrcForm.GBBackground.PaletteID(dx, dy)
                    gBufferBackground.Priority(dx, dy) = gSelection.SrcForm.GBBackground.Priority(dx, dy)
                    gBufferBackground.VRAMEntryAddress(dx, dy) = gSelection.SrcForm.GBBackground.VRAMEntryAddress(dx, dy)
                    gBufferBackground.VRAMEntryBank(dx, dy) = gSelection.SrcForm.GBBackground.VRAMEntryBank(dx, dy)
                    gBufferBackground.XFlip(dx, dy) = gSelection.SrcForm.GBBackground.XFlip(dx, dy)
                    gBufferBackground.YFlip(dx, dy) = gSelection.SrcForm.GBBackground.YFlip(dx, dy)
                    gBufferBackground.BitmapFragmentIndex(dx, dy) = gSelection.SrcForm.GBBackground.BitmapFragmentIndex(dx, dy)
                                
                Next dx
            Next dy
            
            mPasteSel
       
        Case GB_BRUSH
        
            Set gPaletteSrcForm = Nothing
        
            SelectTool GB_BRUSH
            picDisplay_MouseMove Button, Shift, X, Y
        
        Case GB_ZOOM
        
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:picDisplay_MouseUp Error"
End Sub

Private Sub mPasteSel()

    Set gSelection.SrcForm = Me
    
    mGBBackground.Offscreen.Blit mSelBuffer.hdc, 0, 0
       
    If gSelection.SrcForm Is Me Then
        mGBBackground.Offscreen.BlitRaster mSelBuffer.hdc, gSelection.Left * gSelection.CellWidth, gSelection.Top * gSelection.CellHeight, gSelection.SelectionWidth * gSelection.CellWidth, gSelection.SelectionHeight * gSelection.CellHeight, 0, 0, vbDstInvert
    End If
    
    'Update the screen
    mViewport.Draw picDisplay.hdc, mSelBuffer
    picDisplay.Refresh

    mUpdateMore
    
    'If Not mGBBackground Is Nothing Then
    '    mGBBackground.UpdateClients Me
    'End If

End Sub






Private Sub txtBank_Change()

    On Error GoTo HandleErrors

    Dim ret As Long
    Dim indexX As Integer
    Dim indexY As Integer
    
    ret = gSelection.GetFirstElement
    
    Do Until ret < 0
        
        indexX = gSelection.Left + gSelection.CursorX + 1
        indexY = gSelection.Top + gSelection.CursorY + 1
        
        mGBBackground.VRAMEntryBank(indexX, indexY) = val(txtBank.Caption)
        ret = gSelection.GetNextElement
    Loop
    
    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:txtBank_Change Error"
End Sub

Private Sub txtPalID_Change()

    On Error GoTo HandleErrors

    If Not mbNeedApply Then
        Exit Sub
    End If

    Dim ret As Long
    Dim indexX As Integer
    Dim indexY As Integer
    
    ret = gSelection.GetFirstElement
    
    Do Until ret < 0
        
        indexX = gSelection.Left + gSelection.CursorX + 1
        indexY = gSelection.Top + gSelection.CursorY + 1
        
        mGBBackground.PaletteID(indexX, indexY) = val(txtPalID.Text)
        ret = gSelection.GetNextElement
    Loop
    
    mbChanged = True
    mnClipPal = val(txtPalID.Text)
    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:txtPalID_Change Error"
End Sub

Private Sub txtPriority_Change()

    On Error GoTo HandleErrors

    Dim ret As Long
    Dim indexX As Integer
    Dim indexY As Integer
    
    If Not mbNeedApply Then
        Exit Sub
    End If
    
    ret = gSelection.GetFirstElement
    
    Do Until ret < 0
        
        indexX = gSelection.Left + gSelection.CursorX + 1
        indexY = gSelection.Top + gSelection.CursorY + 1
        
        mGBBackground.Priority(indexX, indexY) = val(txtPriority.Text)
        ret = gSelection.GetNextElement
    Loop
    
    mbChanged = True
    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:txtPriority_Change Error"
End Sub

Private Sub vsbScroll_Change()

    On Error GoTo HandleErrors

    If mGBBackground.tBackgroundType = GB_RAWBG Then
        picDisplay.Cls
    End If

    mViewport.ViewportY = vsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBBackground.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBackground:vsbScroll_Change Error"
    
End Sub


Private Sub vsbScroll_Scroll()

    vsbScroll_Change

End Sub
