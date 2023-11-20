VERSION 5.00
Begin VB.Form frmEditSpriteGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmEditSpriteGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   Begin VB.CommandButton cmdBatchPal 
      Caption         =   "Batch &Pal"
      Height          =   360
      Left            =   2760
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CheckBox chkOrigin 
      Caption         =   "&Use Origin"
      Height          =   255
      Left            =   2760
      TabIndex        =   48
      Top             =   4800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdBatchMoveOrigin 
      Caption         =   "&Batch Origin"
      Height          =   360
      Left            =   1440
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdDeleteEntry 
      Caption         =   "R&emove Tile"
      Height          =   360
      Left            =   2760
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CheckBox chkSetOrigin 
      Caption         =   "Move &Origin"
      Height          =   360
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Sa&ve As..."
      Height          =   360
      Left            =   120
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   120
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Frame fraSource 
      Caption         =   "Resource &Files"
      Height          =   1335
      Left            =   4200
      TabIndex        =   25
      Top             =   4800
      Width           =   5415
      Begin VB.CommandButton cmdBrowseVRAM 
         Appearance      =   0  'Flat
         Caption         =   "B&rowse..."
         Height          =   360
         Left            =   1320
         TabIndex        =   29
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdEditVRAM 
         Appearance      =   0  'Flat
         Caption         =   "E&dit..."
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdBrowsePalette 
         Appearance      =   0  'Flat
         Caption         =   "Bro&wse..."
         Height          =   360
         Left            =   3840
         TabIndex        =   33
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdEditPalette 
         Appearance      =   0  'Flat
         Caption         =   "Edi&t..."
         Height          =   360
         Left            =   2640
         TabIndex        =   32
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label txtVRAMSource 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2265
      End
      Begin VB.Label txtPaletteSource 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   31
         Top             =   480
         Width           =   2265
      End
      Begin VB.Label lblVRAMSource 
         AutoSize        =   -1  'True
         Caption         =   "&VRAM:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblPaletteSource 
         AutoSize        =   -1  'True
         Caption         =   "Pa&lette:"
         Height          =   195
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame fraTileEntries 
      Caption         =   "&Tile Info"
      Height          =   2175
      Left            =   4200
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
      Begin VB.CheckBox chkOutline 
         Caption         =   "Outline"
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox txtPriority 
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtBank 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtPalID 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtTileID 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkYFlip 
         Caption         =   "&Y Flip"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   765
         Width           =   735
      End
      Begin VB.CheckBox chkXFlip 
         Caption         =   "&X Flip"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label txtTotalTiles 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   42
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Total:"
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   41
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Priority:"
         Height          =   195
         Index           =   8
         Left            =   1200
         TabIndex        =   39
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Bank:"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Palette ID:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Tile &ID:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "X Of&fset:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Y Offs&et:"
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label txtXOffset 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label txtYOffset 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame fraSpriteEntries 
      Caption         =   "&Sprites"
      Height          =   2415
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdCopySprite 
         Caption         =   "&Copy Sprite"
         Height          =   360
         Left            =   3000
         TabIndex        =   37
         Top             =   1200
         Width           =   2280
      End
      Begin VB.TextBox txtSpriteName 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeleteSprite 
         Caption         =   "&Remove Sprite"
         Height          =   360
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   2280
      End
      Begin VB.CommandButton cmdAddSprite 
         Caption         =   "&Add Sprite"
         Height          =   360
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   2280
      End
      Begin VB.ListBox lstEntries 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "frmEditSpriteGroup.frx":030A
         Left            =   120
         List            =   "frmEditSpriteGroup.frx":030C
         TabIndex        =   1
         Top             =   240
         Width           =   2760
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
         Width           =   465
      End
   End
   Begin VB.Frame fraGridSettings 
      Caption         =   "&Grid Settings"
      Height          =   2175
      Left            =   8040
      TabIndex        =   19
      Top             =   2520
      Width           =   1575
      Begin VB.CheckBox chkUseGrid 
         Caption         =   "&Use Grid"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtGridY 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtGridX 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Grid Y:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Gr&id X:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   3855
      LargeChange     =   8
      Left            =   3840
      Max             =   63
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   255
      LargeChange     =   8
      Left            =   0
      Max             =   63
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3855
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
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frmEditSpriteGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements intResourceClient

Private Const DEF_WIDTH = 9795
Private Const DEF_HEIGHT = 6675

Private msFilename As String
Private mbChanged As Boolean
Private mbCached As Boolean
Private mnSelSprite As Integer
Private mnSelTile As Integer
Private mnGridX As Integer
Private mnGridY As Integer
Private mnLastGX As Integer
Private mnLastGY As Integer
Private mnUseGrid As Integer
Private mbDragOkay As Boolean
Private mbNeedApply As Boolean
Private mnMouseX As Single
Private mnMouseY As Single
Private mbOriginTool As Boolean

'Class variables
Private mGBSpriteGroup As New clsGBSpriteGroup
Private mViewport As New clsViewport

Public Property Get bChanged() As Boolean

    bChanged = mbChanged

End Property

Public Property Let bChanged(bNewValue As Boolean)

    mbChanged = bNewValue

End Property

Public Property Set GBSpriteGroup(oNewValue As clsGBSpriteGroup)

    Set mGBSpriteGroup = oNewValue

End Property

Public Property Get GBSpriteGroup() As clsGBSpriteGroup

    Set GBSpriteGroup = mGBSpriteGroup

End Property



Public Property Set GBVRAM(oNewValue As clsGBVRAM)

    Set mGBSpriteGroup.GBVRAM = oNewValue

End Property

Public Property Get GBVRAM() As clsGBVRAM

    Set GBVRAM = mGBSpriteGroup.GBVRAM

End Property

Private Sub mApply()

    On Error GoTo HandleErrors

    If mnSelTile = 0 Or Not mbNeedApply Then
        Exit Sub
    End If
    
    With mGBSpriteGroup.Sprites(mnSelSprite).Tiles(mnSelTile)
        .XFlip = chkXFlip.value
        .YFlip = chkYFlip.value
        .TileID = val(txtTileID.Text)
        .Bank = val(txtBank.Text)
        .Priority = val(txtPriority.Text)
        .PalID = val(txtPalID.Text)
        .BitmapFragmentIndex = mGBSpriteGroup.GBVRAM.GetBitFragIDFromVRAMAddr((.TileID * 16) + 32768)
    End With
    
    If txtGridX.Text <> "" And txtGridY.Text <> "" Then
        mnGridX = val(txtGridX.Text)
        mnGridY = val(txtGridY.Text)
    End If
    mnUseGrid = chkUseGrid.value
    
    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:mApply Error"

End Sub

Private Sub mBlitTile(tileNum As Integer, Optional Brushing As Boolean)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim j As Integer
        
    i = mnSelSprite
    j = tileNum
    
    If tileNum > mGBSpriteGroup.Sprites(i).TileCount Then
        tileNum = mGBSpriteGroup.Sprites(i).TileCount
        Exit Sub
    End If
    
    If i > 0 And j > 0 Then
        With mGBSpriteGroup.GBVRAM.BitmapFragments(mGBSpriteGroup.Sprites(i).Tiles(j).BitmapFragmentIndex, mGBSpriteGroup.Sprites(i).Tiles(j).Bank)
            If Not .GBBitmap Is Nothing Then
                .GBBitmap.BlitWithPal mGBSpriteGroup.Offscreen.hdc, (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset), 8, 8, CLng(.X), CLng(.Y), mGBSpriteGroup.GBPalette, mGBSpriteGroup.Sprites(i).Tiles(j).PalID, mGBSpriteGroup.Sprites(i).Tiles(j).XFlip, mGBSpriteGroup.Sprites(i).Tiles(j).YFlip, True
            End If
        End With
        If Not Brushing Then
            mGBSpriteGroup.Offscreen.BlitRaster mGBSpriteGroup.Offscreen.hdc, (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset), 8, 8, (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset), vbDstInvert
        End If
    End If

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:mBlitTile Error"
End Sub

Private Sub mDrawLinesOnDisplay()

    On Error GoTo HandleErrors

    Dim i As Long
    Dim X As Long
    Dim Y As Long
        
    If mnGridX > 1 And mnGridY > 1 Then
    
        For Y = 0 To 255 Step mnGridY
            For X = 0 To 255 Step mnGridX
                mGBSpriteGroup.Offscreen.SetPixel X, Y, vbBlack
            Next X
        Next Y
        
    End If
    
    mDrawCrossHairs
    
    mBlitTile mnSelTile

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:mDrawLinesOnDisplay Error"

End Sub

Private Sub mDrawCrossHairs()

    If chkOrigin.value = vbUnchecked Then
        Exit Sub
    End If

    Dim lColor As Long
    
    If mbOriginTool Then
        lColor = vbRed
    Else
        lColor = vbBlack
    End If
    
    mGBSpriteGroup.Offscreen.LineDraw 120, 128, 136, 128, lColor
    mGBSpriteGroup.Offscreen.LineDraw 128, 120, 128, 136, lColor
    
    mBlitTile mnSelTile

End Sub




Private Sub mResetOrigin(newX As Integer, newY As Integer)

    Dim i As Integer
    
    With mGBSpriteGroup.Sprites(mnSelSprite)
        For i = 1 To .TileCount
            .Tiles(i).XOffset = .Tiles(i).XOffset + newX
            .Tiles(i).yOffset = .Tiles(i).yOffset + newY
        Next i
    End With

    mbChanged = True

End Sub

Private Sub mUpdateMore()

    On Error GoTo HandleErrors

'Refresh the data to be displayed
    txtSpriteName.Text = mGBSpriteGroup.Sprites(mnSelSprite).Name
    
    If mGBSpriteGroup.SpriteCount = 0 Then
        Exit Sub
    End If
    
    mbNeedApply = False
    
    With mGBSpriteGroup.Sprites(mnSelSprite).Tiles(mnSelTile)
        txtXOffset.Caption = CStr(.XOffset) - 128
        txtYOffset.Caption = CStr(.yOffset) - 128
        chkXFlip.value = .XFlip
        chkYFlip.value = .YFlip
        txtTileID.Text = CStr(.TileID)
        txtBank.Text = CStr(.Bank)
        txtPriority.Text = CStr(.Priority)
        txtPalID.Text = CStr(.PalID)
        txtGridX.Text = CStr(mnGridX)
        txtGridY.Text = CStr(mnGridY)
        chkUseGrid.value = mnUseGrid
    End With
    
    mbNeedApply = True
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, " frmEditSpriteGroup:mUpdateMore Error"
End Sub

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Private Sub chkOrigin_Click()

    intResourceClient_Update

End Sub

Private Sub chkOutline_Click()

    intResourceClient_Update

End Sub

Private Sub chkSetOrigin_Click()

    If chkSetOrigin.value = vbChecked Then
        mbOriginTool = True
        Screen.MousePointer = vbCrosshair
    Else
        mbOriginTool = False
        Screen.MousePointer = 0
    End If
        
    intResourceClient_Update

End Sub

Private Sub chkUseGrid_Click()

    txtGridX.Enabled = chkUseGrid.value
    txtGridY.Enabled = chkUseGrid.value
    mnUseGrid = chkUseGrid.value
    
    If chkUseGrid.value = False Then
        mnLastGX = mnGridX
        mnLastGY = mnGridY
        mnGridX = 1
        mnGridY = 1
    Else
        mnGridX = mnLastGX
        mnGridY = mnLastGY
    End If
    
    mUpdateMore
    intResourceClient_Update
    
    If Me.Visible = True Then
        picDisplay.SetFocus
    End If
    
End Sub

Private Sub chkXFlip_Click()

    mApply

End Sub

Private Sub chkYFlip_Click()

    mApply

End Sub


Private Sub cmdAddSprite_Click()

    On Error GoTo HandleErrors

    Dim sName As String
    sName = InputBox("Enter frame name:", "Input")
    
    If sName = "" Then
        Exit Sub
    End If
    
    mGBSpriteGroup.AddSprite
    mGBSpriteGroup.Sprites(mGBSpriteGroup.SpriteCount).Name = UCase$(Replace(sName, " ", "_"))
    
    mnSelTile = 0
    mnSelSprite = mGBSpriteGroup.SpriteCount
    
    mPopulateSpriteEntryList lstEntries.ListCount

    intResourceClient_Update
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdAddSprite_Click Error"

End Sub

Private Sub cmdBatchMoveOrigin_Click()

    Dim i As Integer
    Dim startIndex As Integer
    Dim dx As Integer
    Dim dy As Integer
    
    dx = val(InputBox("Input delta x:", "Data Input"))
    dy = val(InputBox("Input delta y:", "Data Input"))
    
    startIndex = lstEntries.ListIndex
    
    For i = 0 To lstEntries.ListCount - 1
        lstEntries.ListIndex = i
        mResetOrigin -dx, -dy
    Next i
    
    lstEntries.ListIndex = startIndex
    
    intResourceClient_Update

End Sub

Private Sub cmdBatchPal_Click()

    Dim i As Integer
    Dim dummy As String
    
    dummy = InputBox("Input palette ID:", "Batch Palette Change")
    If dummy = "" Then
        Exit Sub
    End If
    
    For i = 0 To mGBSpriteGroup.Sprites(mnSelSprite).TileCount
        mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).PalID = val(dummy)
    Next i

    intResourceClient_Update

End Sub

Private Sub cmdBrowsePalette_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Get the filename for the palette
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    With mdiMain.Dialog
        .InitDir = mGBSpriteGroup.intResource_ParentPath & "\Palettes"
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
        Set mGBSpriteGroup.GBPalette = gResourceCache.GetResourceFromFile(sFilename, mGBSpriteGroup)
    
        mGBSpriteGroup.sPaletteFile = sFilename
    
    End With

    intResourceClient_Update

Exit Sub

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sFilename & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = mGBSpriteGroup.intResource_ParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                Exit Sub
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, mGBSpriteGroup.intResource_ParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, mGBSpriteGroup.intResource_ParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
                        
            Resume
            
        End With
    
    Else
        MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdBrowsePalette_Click Error"
    End If
End Sub

Private Sub cmdBrowseVRAM_Click()

'***************************************************************************
'   Get the filename for the VRAM
'***************************************************************************
    
    With mdiMain.Dialog
        .InitDir = mGBSpriteGroup.intResource_ParentPath & "\VRAMs"
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
        Set mGBSpriteGroup.GBVRAM = gResourceCache.GetResourceFromFile(sFilename, mGBSpriteGroup)
        gbOpeningChild = False
    
        mGBSpriteGroup.sVRAMFile = sFilename
    
    End With
    
    intResourceClient_Update

Exit Sub

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sFilename & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = mGBSpriteGroup.intResource_ParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                Exit Sub
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, mGBSpriteGroup.intResource_ParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, mGBSpriteGroup.intResource_ParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
                        
            Resume
            
        End With
    
    Else

        MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdBrowseVRAM_Click Error"
    End If
End Sub

Private Sub cmdCopySprite_Click()

    On Error GoTo HandleErrors

    Dim i As Integer
    
    mGBSpriteGroup.AddSprite
    mGBSpriteGroup.Sprites(mGBSpriteGroup.SpriteCount).Name = mGBSpriteGroup.Sprites(mnSelSprite).Name & "_COPY"
    
    For i = 1 To mGBSpriteGroup.Sprites(mnSelSprite).TileCount
        mGBSpriteGroup.Sprites(mGBSpriteGroup.SpriteCount).AddTile
        With mGBSpriteGroup.Sprites(mGBSpriteGroup.SpriteCount).Tiles(i)
            .XOffset = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).XOffset
            .yOffset = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).yOffset
            .XFlip = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).XFlip
            .YFlip = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).YFlip
            .Bank = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).Bank
            .BitmapFragmentIndex = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).BitmapFragmentIndex
            .PalID = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).PalID
            .Priority = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).Priority
            .TileID = mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).TileID
        End With
    Next i
    
    mnSelTile = 0
    mnSelSprite = mGBSpriteGroup.SpriteCount
    
    mPopulateSpriteEntryList lstEntries.ListCount

    intResourceClient_Update
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdAddSprite_Click Error"
End Sub

Private Sub cmdDeleteEntry_Click()

    On Error GoTo HandleErrors

'Check for errors
    If mGBSpriteGroup.Sprites(mnSelSprite).TileCount <= 0 Or mnSelTile = 0 Then
        Exit Sub
    End If
    
'Delete the appropriate entry
    mGBSpriteGroup.Sprites(mnSelSprite).DeleteTile mnSelTile
    
    If mGBSpriteGroup.Sprites(mnSelSprite).TileCount <= 0 Then
        mnSelTile = 0
    End If
    
'Update the visual
    mbChanged = True
    
    intResourceClient_Update

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdDeleteEntry_Click Error"
End Sub

Private Sub cmdDeleteSprite_Click()

    On Error GoTo HandleErrors

    Dim curSel As Integer
    
'Check for errors
    If lstEntries.ListCount <= 0 Or lstEntries.ListIndex = -1 Or mnSelSprite = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the currently selected sprite and all of its tiles?  This operation cannot be undone!", vbQuestion Or vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
'Delete the appropriate entry
    curSel = lstEntries.ListIndex

    mGBSpriteGroup.DeleteSprite mnSelSprite
    
    If mGBSpriteGroup.SpriteCount = 0 Then
        mnSelSprite = 0
        mnSelTile = 0
    End If
    
    mPopulateSpriteEntryList curSel
    mUpdateMore

'Update the visual
    mbChanged = True
    
    intResourceClient_Update

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdDeleteSprite_Click Error"
End Sub

Private Sub cmdEditPalette_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Open the palette file in an editor
'***************************************************************************

    On Error GoTo HandleErrors

    Screen.MousePointer = vbHourglass

    If mGBSpriteGroup.GBPalette Is Nothing Then
        MsgBox "No Palette file specified!", vbCritical, "Error"
        Exit Sub
    End If
    
    Dim frm As New frmEditPalette
    gResourceCache.AddResourceToCache mGBSpriteGroup.sPaletteFile, mGBSpriteGroup.GBPalette, frm
    Set frm.GBPalette = mGBSpriteGroup.GBPalette
    frm.sFilename = mGBSpriteGroup.intResource_ParentPath & "\Palettes\" & GetTruncFilename(mGBSpriteGroup.sPaletteFile)
    frm.Show

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdEditPalette_Click"
    Screen.MousePointer = 0
End Sub

Private Sub cmdEditVRAM_Click()

'***************************************************************************
'   Open the VRAM file in an editor
'***************************************************************************

    On Error GoTo HandleErrors
    
    Screen.MousePointer = vbHourglass

    If mGBSpriteGroup.GBVRAM Is Nothing Then
        MsgBox "No VRAM file specified!", vbCritical, "Error"
        Exit Sub
    End If
    
    Dim frm As New frmEditVRAM
    gResourceCache.AddResourceToCache mGBSpriteGroup.sVRAMFile, mGBSpriteGroup.GBVRAM, frm
    Set frm.GBVRAM = mGBSpriteGroup.GBVRAM
    frm.sFilename = mGBSpriteGroup.intResource_ParentPath & "\VRAMs\" & GetTruncFilename(mGBSpriteGroup.sVRAMFile)
    frm.Show

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdEditVRAM_Click Error"
    Screen.MousePointer = 0
End Sub



Private Sub cmdSave_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current sprite group into a .spr file
'***************************************************************************

    PackFile mGBSpriteGroup.intResource_ParentPath & "\Sprites\" & GetTruncFilename(msFilename), mGBSpriteGroup
    mbChanged = False

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdSave_Click Error"
End Sub

Private Sub cmdSaveAs_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current sprite group under a new filename
'***************************************************************************

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Sprite Group"
        .Filename = ""
        .Filter = "GB Sprite Groups (*.spr)|*.spr"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBSpriteGroup
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBSpriteGroup.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBSpriteGroup, Me
        
        mGBSpriteGroup.Offscreen.Create 256, 256
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:cmdSaveAs_Click Error"
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete And mnSelTile <> 0 Then
        cmdDeleteEntry_Click
    End If
    
    With mGBSpriteGroup.Sprites(mnSelSprite).Tiles(mnSelTile)
            
        If KeyCode = vbKeyLeft Then
            If mbOriginTool Then
                mResetOrigin -1, 0
                intResourceClient_Update
                picDisplay.SetFocus
            ElseIf mnSelTile <> 0 Then
                .XOffset = .XOffset - 1
                intResourceClient_Update
                picDisplay.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyCode = vbKeyRight Then
            If mbOriginTool Then
                mResetOrigin 1, 0
                intResourceClient_Update
                picDisplay.SetFocus
            ElseIf mnSelTile <> 0 Then
                .XOffset = .XOffset + 1
                intResourceClient_Update
                picDisplay.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyCode = vbKeyUp Then
            If mbOriginTool Then
                mResetOrigin 0, -1
                intResourceClient_Update
                picDisplay.SetFocus
            ElseIf mnSelTile <> 0 Then
                .yOffset = .yOffset - 1
                intResourceClient_Update
                picDisplay.SetFocus
            End If
            Exit Sub
        End If
            
        If KeyCode = vbKeyDown Then
            If mbOriginTool Then
                mResetOrigin 0, 1
                intResourceClient_Update
                picDisplay.SetFocus
            ElseIf mnSelTile <> 0 Then
                .yOffset = .yOffset + 1
                intResourceClient_Update
                picDisplay.SetFocus
            End If
            Exit Sub
        End If
        
    End With

End Sub



Private Sub Form_Load()

    On Error GoTo HandleErrors

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT

    mnLastGX = 8
    mnLastGY = 8
    mnGridX = 8
    mnGridY = 8
    mnUseGrid = 1
    
'Organize forms on screen
    CleanUpForms Me
        
    mnSelSprite = 0
        
'Set up visual objects
    mGBSpriteGroup.Offscreen.Create 256, 256
    
    With mViewport
        .Zoom = 1
        
        .ViewportWidth = 256
        .ViewportHeight = 256
        
        .SourceWidth = 256
        .SourceHeight = 256
        
        .DisplayWidth = 256
        .DisplayHeight = 256
        
        hsbScroll.Max = .hScrollMax
        vsbScroll.Max = .vScrollMax
    End With
    
    mPopulateSpriteEntryList 0
    mUpdateMore
    
    intResourceClient_Update

    SelectTool GB_POINTER

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:Form_Load Error"

End Sub

Private Sub mPopulateSpriteEntryList(nSelection As Integer)

    On Error GoTo HandleErrors

'***************************************************************************
'   Place info on each sprite group entry into the list box
'***************************************************************************

'Clear the list box before we begin
    lstEntries.Clear
    
    Dim i As Integer
    Dim j As Integer

'Add the data to the list box with AddItem
    For i = 1 To mGBSpriteGroup.SpriteCount
        lstEntries.AddItem mGBSpriteGroup.Sprites(i).Name
    Next i
    
'Error checking
    If nSelection > lstEntries.ListCount - 1 Then
        nSelection = lstEntries.ListCount - 1
    End If
    If nSelection < 0 Then
        nSelection = 0
    End If
    If lstEntries.ListCount <= 0 Then
        nSelection = -1
    End If
    
'Select the element specified by nSelection
    lstEntries.ListIndex = nSelection

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:mPopulateSpriteEntryList Error"
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

    On Error GoTo HandleErrors
    
    
    gResourceCache.ReleaseClient Me
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:Form_Unload Error"
End Sub

Private Sub hsbScroll_Change()

    On Error GoTo HandleErrors

    mViewport.ViewportX = hsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBSpriteGroup.Offscreen
    picDisplay.Refresh

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:hsbScroll_Change Error"

End Sub

Private Sub hsbScroll_GotFocus()

    Me.SetFocus

End Sub


Private Sub hsbScroll_Scroll()

    hsbScroll_Change

End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

'***************************************************************************
'   Update the visual display
'***************************************************************************

'Set up bitmap fragments
    mGBSpriteGroup.GBVRAM.EnumBitmapFragments
    mGBSpriteGroup.Offscreen.Cls
    
'Draw grid on the display
    mDrawLinesOnDisplay
    
'Draw bitmap fragments
    If Not mGBSpriteGroup.GBVRAM Is Nothing Then
        
        Dim i As Integer
        Dim j As Integer
            
        If mGBSpriteGroup.SpriteCount <> 0 Then
            i = mnSelSprite
            For j = 1 To mGBSpriteGroup.Sprites(i).TileCount
                With mGBSpriteGroup.GBVRAM.BitmapFragments(mGBSpriteGroup.GBVRAM.GetBitFragIDFromVRAMAddr(((mGBSpriteGroup.Sprites(i).Tiles(j).TileID * 16) + 32768)), mGBSpriteGroup.Sprites(i).Tiles(j).Bank)
                    If Not .GBBitmap Is Nothing Then
                        .GBBitmap.BlitWithPal mGBSpriteGroup.Offscreen.hdc, (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset), 8, 8, CLng(.X), CLng(.Y), mGBSpriteGroup.GBPalette, mGBSpriteGroup.Sprites(i).Tiles(j).PalID, mGBSpriteGroup.Sprites(i).Tiles(j).XFlip, mGBSpriteGroup.Sprites(i).Tiles(j).YFlip, True
                    End If
                    If chkOutline.value = vbChecked Then
                        mGBSpriteGroup.Offscreen.LineDraw (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset), (mGBSpriteGroup.Sprites(i).Tiles(j).XOffset) + 8, (mGBSpriteGroup.Sprites(i).Tiles(j).yOffset) + 8, vbRed, True
                    End If
                End With
            Next j
        End If
        
    End If
    
'Update the screen
    mDrawCrossHairs
    mViewport.Draw picDisplay.hdc, mGBSpriteGroup.Offscreen

    mUpdateMore
    Me.Caption = GetTruncFilename(msFilename)
    txtPaletteSource.Caption = GetTruncFilename(mGBSpriteGroup.sPaletteFile)
    txtVRAMSource.Caption = GetTruncFilename(mGBSpriteGroup.sVRAMFile)
    
    txtTotalTiles.Caption = CStr(mGBSpriteGroup.Sprites(mnSelSprite).TileCount)
    
    picDisplay.Refresh
    
    If tType = GB_ACTIVEEDITOR Then
        If Not mGBSpriteGroup Is Nothing Then
            mGBSpriteGroup.UpdateClients Me
            mGBSpriteGroup.UpdateClients Me
        End If
    End If

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:intResourceClient_Update Error"
End Sub


Private Sub lstEntries_Click()

    mnSelSprite = lstEntries.ListIndex + 1
    mnSelTile = 0
    
    intResourceClient_Update

End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors
    Dim click As New clsPoint

    If mbOriginTool Then
        Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 1, 1)
        mResetOrigin 128 - click.X, 128 - click.Y
        
        intResourceClient_Update
        
        Exit Sub
    End If

    mnMouseX = X
    mnMouseY = Y
    
    mUpdateMore

    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), mnGridX, mnGridY)
    
    Select Case gnTool
        
        Case GB_POINTER
            
            Dim i As Integer
            Dim dx As Integer
            Dim dy As Integer
            Dim cx As Integer
            Dim cy As Integer
            
            Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 1, 1)
            
            For i = mGBSpriteGroup.Sprites(mnSelSprite).TileCount To 1 Step -1
                If i < 1 Then
                    Exit For
                End If
                
                With mGBSpriteGroup.Sprites(mnSelSprite)
                    dx = .Tiles(i).XOffset
                    dy = .Tiles(i).yOffset
                    cx = click.X
                    cy = click.Y
                    If (cx >= dx And cx < dx + 8) And (cy >= dy And cy < dy + 8) Then
                        
                        mbDragOkay = True
                        mnSelTile = i
                        intResourceClient_Update
                        Exit Sub
                    End If
                End With
            Next i
            
            mbDragOkay = True
            mnSelTile = 0
            
        Case GB_MARQUEE
            
        Case GB_BRUSH
            
            If gSelection.SrcForm Is Nothing Then
                SelectTool GB_POINTER
                Exit Sub
            End If
            
            If Not TypeOf gSelection.SrcForm Is frmEditVRAM Then
                SelectTool GB_POINTER
                picDisplay_MouseDown Button, Shift, X, Y
                Exit Sub
            End If
            
            If mGBSpriteGroup.SpriteCount = 0 Then
                MsgBox "You must add a sprite first!", vbInformation, "Information"
                mnSelSprite = 0
                Exit Sub
            End If
            
            Dim ret As Integer
            Dim xIndex As Integer
            Dim yIndex As Integer
                        
        'Transfer data from VRAM to sprite group
            
            Set mGBSpriteGroup.GBVRAM = gSelection.SrcForm.GBVRAM
            
            ret = gSelection.GetFirstElement
            
            Do Until ret < 0
                xIndex = (click.X + gSelection.CursorX) * mnGridX
                yIndex = (click.Y + gSelection.CursorY) * mnGridY
                If xIndex < 0 Or xIndex > 256 Or yIndex < 0 Or yIndex > 256 Or ret > 256 Then
                Else
                    With mGBSpriteGroup.Sprites(mnSelSprite)
                        Dim j As Integer
                        For i = 1 To .TileCount
                            If .Tiles(i).XOffset = xIndex And .Tiles(i).yOffset = yIndex Then
                                .DeleteTile i
                                Exit For
                            End If
                        Next i
                        .AddTile
                        .Tiles(.TileCount).TileID = ret - 1
                        .Tiles(.TileCount).Bank = gSelection.SrcForm.SelectedBank
                        .Tiles(.TileCount).XOffset = xIndex
                        .Tiles(.TileCount).yOffset = yIndex
                        .Tiles(.TileCount).BitmapFragmentIndex = mGBSpriteGroup.GBVRAM.GetBitFragIDFromVRAMAddr((.Tiles(.TileCount).TileID * 16) + 32768)
                    
                        mBlitTile mGBSpriteGroup.Sprites(mnSelSprite).TileCount, True
                        
                    End With
                End If
                ret = gSelection.GetNextElement
            Loop
            
        'Update VRAM Source text box
            If txtVRAMSource.Caption = "" Then
                If gSelection.SrcForm.sFilename <> "" Then
                    mGBSpriteGroup.sVRAMFile = gSelection.SrcForm.sFilename
                End If
            End If
        
        'Set variable used for saving
            mbChanged = True
            
        'Update visual
            mViewport.Draw picDisplay.hdc, mGBSpriteGroup.Offscreen
    
        'Draw grid on the display
            mDrawLinesOnDisplay
            
            mPopulateSpriteEntryList mnSelSprite - 1
            
            picDisplay.Refresh
            
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
    
            If gSelection.SrcForm Is Nothing Then
                Exit Sub
            End If
            
            If Not TypeOf gSelection.SrcForm Is frmEditPalette Then
                Exit Sub
            End If
            
            Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 1, 1)
            
            For i = mGBSpriteGroup.Sprites(mnSelSprite).TileCount To 1 Step -1
                If i < 1 Then
                    Exit For
                End If
                
                With mGBSpriteGroup.Sprites(mnSelSprite)
                    dx = .Tiles(i).XOffset
                    dy = .Tiles(i).yOffset
                    cx = click.X
                    cy = click.Y
                    If (cx >= dx And cx < dx + 8) And (cy >= dy And cy < dy + 8) Then
                        
                        mbDragOkay = True
                        mGBSpriteGroup.Sprites(mnSelSprite).Tiles(i).PalID = gSelection.SrcForm.SelectedPalette
                        intResourceClient_Update
                        Exit Sub
                    End If
                End With
            Next i
            
            mbDragOkay = True
            mnSelTile = 0
    
    End Select

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup Error"
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    On Error GoTo HandleErrors
 
    If mnMouseX = X And mnMouseY = Y Then
        Exit Sub
    End If
 
    If X >= picDisplay.width Or X < 0 Or Y >= picDisplay.height Or Y < 0 Then
        Exit Sub
    End If
    
    mnMouseX = X
    mnMouseY = Y
    
    mUpdateMore
     
    Dim click As New clsPoint

    Select Case gnTool
        
        Case GB_POINTER
            
            picDisplay.MousePointer = vbArrow
            
            If Button = vbLeftButton And mbDragOkay Then
                
                Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), mnGridX, mnGridY)
                mGBSpriteGroup.Sprites(mnSelSprite).Tiles(mnSelTile).XOffset = click.X * mnGridX
                mGBSpriteGroup.Sprites(mnSelSprite).Tiles(mnSelTile).yOffset = click.Y * mnGridY

                mbChanged = True
            
                intResourceClient_Update
                            
            End If
            
        Case GB_MARQUEE
        
            picDisplay.MousePointer = vbArrow
        
        Case GB_BRUSH
        
            picDisplay.MouseIcon = mdiMain.picDragAdd
            picDisplay.MousePointer = vbCustom
                
        Case GB_ZOOM
        
            picDisplay.MouseIcon = mdiMain.picZoom
            picDisplay.MousePointer = vbCustom
        
        Case GB_SETTER
        
            picDisplay.MousePointer = vbUpArrow
        
    End Select

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:picDisplay_MouseMove Error"
End Sub








Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mnMouseX = X
    mnMouseY = Y
    
    mUpdateMore

    If gnTool = GB_POINTER And Button = vbLeftButton Then
        mbDragOkay = False
        
        If mnSelTile = 0 Then
            intResourceClient_Update
        End If
        
    End If

End Sub





Private Sub txtBank_Change()
        
    mApply

End Sub

Private Sub txtGridX_Change()
        
    mApply

End Sub

Private Sub txtGridY_Change()
        
    mApply

End Sub


Private Sub txtPalID_Change()
        
    mApply

End Sub





Private Sub txtPriority_Change()

    mApply

End Sub


Private Sub txtSpriteName_Change()
        
    Dim nStart As Integer
    nStart = txtSpriteName.SelStart
        
    mGBSpriteGroup.Sprites(mnSelSprite).Name = UCase$(Replace(txtSpriteName.Text, " ", "_"))
    mPopulateSpriteEntryList mnSelSprite - 1
    mbChanged = True
    txtSpriteName.SelStart = nStart

End Sub


Private Sub txtTileID_Change()
        
    mApply

End Sub

Private Sub vsbScroll_Change()

    On Error GoTo HandleErrors

    mViewport.ViewportY = vsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBSpriteGroup.Offscreen
    picDisplay.Refresh

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditSpriteGroup:vsbScroll_Change Error"

End Sub

Private Sub vsbScroll_GotFocus()

    Me.SetFocus

End Sub


Private Sub vsbScroll_Scroll()

    vsbScroll_Change

End Sub


