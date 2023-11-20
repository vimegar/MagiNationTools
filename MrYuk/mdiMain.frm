VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Mr. Yuk 2000"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12225
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResourceCacheView 
      Align           =   4  'Align Right
      Height          =   9135
      Left            =   9300
      ScaleHeight     =   605
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2925
      Begin VB.ListBox lstResources 
         Height          =   5130
         ItemData        =   "mdiMain.frx":030A
         Left            =   120
         List            =   "mdiMain.frx":030C
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Resources In Cache:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1500
      End
   End
   Begin VB.PictureBox picDragAdd 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "mdiMain.frx":030E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   811
      TabIndex        =   1
      Top             =   9135
      Visible         =   0   'False
      Width           =   12225
      Begin VB.PictureBox picZoom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2160
         Picture         =   "mdiMain.frx":0460
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picDropper 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1440
         Picture         =   "mdiMain.frx":05B2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picSetter 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   420
         Left            =   960
         Picture         =   "mdiMain.frx":0704
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.PictureBox picBucket 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   420
         Left            =   480
         Picture         =   "mdiMain.frx":0DA8
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Timer tmrUndo 
      Interval        =   1000
      Left            =   600
      Top             =   7920
   End
   Begin MSComctlLib.ImageList imgTools 
      Left            =   600
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":14AA
            Key             =   "Pointer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1BBE
            Key             =   "Brush"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":22D2
            Key             =   "Marquee"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29E6
            Key             =   "Zoom"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":30FA
            Key             =   "Setter"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":380E
            Key             =   "Bucket"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3F22
            Key             =   "Replace"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   3  'Align Left
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   16113
      ButtonWidth     =   820
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgTools"
      DisabledImageList=   "imgTools"
      HotImageList    =   "imgTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pointer"
            Object.ToolTipText     =   "Pointer"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Brush"
            Object.ToolTipText     =   "Brush"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Marquee"
            Object.ToolTipText     =   "Marquee"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoom"
            Object.ToolTipText     =   "Zoom"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Setter"
            Object.ToolTipText     =   "Setter"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bucket"
            Object.ToolTipText     =   "Bucket"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Object.ToolTipText     =   "Replace"
            ImageIndex      =   7
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrHotKeys 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   7920
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1200
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Begin VB.Menu mnuFileNewBitmap 
            Caption         =   "&Bitmap..."
         End
         Begin VB.Menu mnuFileNewVRAM 
            Caption         =   "&VRAM..."
         End
         Begin VB.Menu mnuFileNewPalette 
            Caption         =   "Pa&lette..."
         End
         Begin VB.Menu mnuFileNewMap 
            Caption         =   "&Map..."
         End
         Begin VB.Menu mnuFileNewBackgroundsPattern 
            Caption         =   "&Pattern..."
         End
         Begin VB.Menu mnuFileNewCollision 
            Caption         =   "&Collision Map..."
         End
         Begin VB.Menu mnuFileNewBackgroundsRawBG 
            Caption         =   "&Raw BG..."
         End
         Begin VB.Menu mnuFileNewSpriteGroup 
            Caption         =   "&Sprite Group..."
         End
      End
      Begin VB.Menu mnuNewProjectStructure 
         Caption         =   "New Project Structure..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenNormal 
         Caption         =   "&Open File..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open File &Batch..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileOpenAdvancedBitmap 
         Caption         =   "Open &Advanced Bitmap..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuBinaryOutput 
      Caption         =   "&Tools"
      Begin VB.Menu mnuColCodesOpen 
         Caption         =   "&Collision Codes Tool"
      End
      Begin VB.Menu mnuToolsAdvancedBitmapTool 
         Caption         =   "&Advanced Bitmap Tool..."
      End
   End
   Begin VB.Menu mnuToolsPack 
      Caption         =   "&Output"
      Begin VB.Menu mnuBinaryOutputPackToBin 
         Caption         =   "Export to &CGB File..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuToolsPackRLE 
         Caption         =   "&RLE Compression..."
      End
      Begin VB.Menu mnuGBOuputBankPacker 
         Caption         =   "&Bank Packer..."
      End
      Begin VB.Menu mnuToolsOutputMapPictureOutput 
         Caption         =   "&Map Picture Output..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewCache 
         Caption         =   "Resource &Cache"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Mr. Yuk 2000..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'   VB Interface Setup
'***************************************************************************

    Option Explicit

'***************************************************************************
'   Form dimensions
'***************************************************************************

    Private Const DEF_WIDTH = 12345
    Private Const DEF_HEIGHT = 10035



Private Sub MDIForm_Load()
    
'***************************************************************************
'   Resize form and initialize global handle
'***************************************************************************
    
    On Error GoTo HandleErrors
    
'Set form's size
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
'Show form
    Me.Show
    
'Initialize handle
    glMainHDC = GetDC(Me.hwnd)
    
'Set current path for dialog boxes
    gsCurPath = App.Path
    
    tmrHotKeys.Enabled = True
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:MDIForm_Load Error"
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'Standard mouse-down logic (based on selected tool)
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
        
        Case GB_MARQUEE
        
        Case GB_BRUSH
            
        Case GB_ZOOM
    
    End Select

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'Standard mouse-move logic (based on selected tool)
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
        
        'Set cursor to "arrow"
            
            Me.MousePointer = vbArrow
        
        Case GB_MARQUEE
        
        'Set cursor to "arrow"
            Me.MousePointer = vbArrow
        
        Case GB_BRUSH
            
            Me.MousePointer = vbArrow
    
        Case GB_ZOOM
    
            Me.MousePointer = vbArrow
    
    End Select

End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'***************************************************************************
'   Prepare for closing of program
'***************************************************************************

    On Error GoTo HandleErrors

'Set global flag to let other forms know to close properly
    gbClosingApp = True

'Unload the tool window
    'Unload frmMapTools

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:MDIForm_QueryUnload Error"

End Sub


Private Sub MDIForm_Resize()

    On Error GoTo HandleErrors

'***************************************************************************
'   Keep the tool window consistent with the main window
'***************************************************************************
    
    'If Me.WindowState = vbMinimized Then
    '    frmMapTools.WindowState = vbMinimized
    'Else
    '    frmMapTools.WindowState = vbNormal
    'End If

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:MDIForm_Resize Error"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    On Error GoTo HandleErrors

'***************************************************************************
'   Release memory as the program closes
'***************************************************************************

    Dim i As Integer
    
    For i = 0 To Forms.count - 1
        If Not (Forms(i) Is Me) Then
            Unload Forms(i)
        End If
    Next i

    DeleteDC glMainHDC
    
    gResourceCache.CloseCache
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:MDIForm_Unload Error"
End Sub


Private Sub mnuBinaryOutput_Click()

    SelectTool GB_POINTER

End Sub

Private Sub mnuBinaryOutputPackToBin_Click()

    On Error GoTo HandleErrors

    frmPackToBin.Show vbModal
    Me.Show

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuBinaryOuputPackToBin Error"
End Sub

Private Sub mnuColCodesOpen_Click()

    gFrmColCodes.Show
    gFrmColCodes.tmrHide.Enabled = True

End Sub

Private Sub mnuFile_Click()

    SelectTool GB_POINTER

End Sub

Private Sub mnuFileExit_Click()
    
'***************************************************************************
'End the program and unload all forms
'***************************************************************************
    
    Unload Me
    End
    
End Sub

Private Sub mnuFileNewBackgroundsPattern_Click()

'***************************************************************************
'   Create new GB pattern and open it in an editor
'***************************************************************************
 
    On Error GoTo HandleErrors
   
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "pat"
        .DialogTitle = "Create New GB Pattern - Choose Filename"
        .Filename = ""
        .Filter = "GB Patterns (*.pat)|*.pat"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditBackground
        Dim res As New clsGBBackground
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBBackground = res
        
        frm.bOpening = True
        frm.bChanged = True
        frm.sFilename = .Filename
        frm.GBBackground.tBackgroundType = GB_PATTERNBG
        frm.GBBackground.nWidth = 32
        frm.GBBackground.nHeight = 32
        frm.GBBackground.intResource_ParentPath = sParentPath
        
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewBackgroundsPattern_Click Error"
End Sub

Private Sub mnuFileNewBackgroundsRawBG_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB raw bg and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "bg"
        .DialogTitle = "Create New GB BG - Choose Filename"
        .Filename = ""
        .Filter = "GB BGs (*.bg)|*.bg"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditBackground
        Dim res As New clsGBBackground
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBBackground = res
        
        frm.bChanged = True
        frm.GBBackground.tBackgroundType = GB_RAWBG
        frm.sFilename = .Filename
        frm.GBBackground.intResource_ParentPath = sParentPath
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With

Exit Sub
                                 
HandleErrors:
    If Err.Description <> "Object was unloaded" Then
        MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewBackgroundsRawBG_Click Error"
    End If
End Sub

Private Sub mnuFileNewBitmap_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB bitmap and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "bit"
        .DialogTitle = "Create New GB Bitmap - Choose Filename"
        .Filename = ""
        .Filter = "GB Bitmaps (*.bit)|*.bit"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
       
    'Create new editor
        Dim frm As New frmEditBitmap
        Dim res As New clsGBBitmap
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBBitmap = res
        frm.bChanged = True
        frm.sFilename = .Filename
        frm.GBBitmap.intResource_ParentPath = sParentPath '& "Bitmaps\"
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With
    
 Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewBitmap_Click Error"
End Sub

Private Sub mnuFileNewCollision_Click()
    
    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB collision map and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "clm"
        .DialogTitle = "Create New GB Collision Map - Choose Filename"
        .Filename = ""
        .Filter = "GB Collision Maps (*.clm)|*.clm"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditCollisionMap
        Dim File As String
        Dim res As New clsGBCollisionMap
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set res.GBCollisionCodes = gGBCollisionCodes
        Set frm.GBCollisionMap = res
        
        File = .Filename
        frm.sFilename = File
        frm.GBCollisionMap.intResource_ParentPath = sParentPath
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With

Exit Sub
                                 
HandleErrors:
    If Err.Description <> "Object was unloaded" Then
        MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewCollisionMap_Click Error"
    End If
End Sub

Private Sub mnuFileNewMap_Click()

'***************************************************************************
'   Create new GB map and open it in an editor
'***************************************************************************
  
    On Error GoTo HandleErrors
  
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "map"
        .DialogTitle = "Create New GB Map - Choose Filename"
        .Filename = ""
        .Filter = "GB Maps (*.map)|*.map"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditMap
        Dim File As String
        Dim res As New clsGBMap
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBMap = res
        
        File = .Filename
        frm.GBMap.GBBackground.tBackgroundType = GB_PATTERNBG
        frm.sFilename = File
        frm.GBMap.intResource_ParentPath = sParentPath
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With
    
Exit Sub
                                 
HandleErrors:
    If Err.Description <> "Object was unloaded" Then
        MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewMap_Click Error"
    End If
End Sub

Private Sub mnuFileNewPalette_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB palette and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "pal"
        .DialogTitle = "Create New GB Palette - Choose Filename"
        .Filename = ""
        .Filter = "GB Palettes (*.pal)|*.pal"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditPalette
        Dim res As New clsGBPalette
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBPalette = res
        frm.bChanged = True
        frm.sFilename = .Filename
        frm.GBPalette.intResource_ParentPath = sParentPath
        frm.Show
         
    'Organize forms
        CleanUpForms Me
        
    End With
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewPalette_Click Error"
End Sub


Private Sub mnuFileNewSpriteGroup_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB sprite group and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "spr"
        .DialogTitle = "Create New GB Sprite Group - Choose Filename"
        .Filename = ""
        .Filter = "GB Sprite Groups (*.spr)|*.spr"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditSpriteGroup
        Dim File As String
        Dim res As New clsGBSpriteGroup
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBSpriteGroup = res
        File = .Filename
        frm.bChanged = True
        frm.sFilename = File
        frm.GBSpriteGroup.intResource_ParentPath = sParentPath
        frm.Show
        
    'Organize forms
        CleanUpForms Me
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewSpriteGroup_Click Error"

End Sub

Private Sub mnuFileNewVRAM_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create new GB VRAM and open it in an editor
'***************************************************************************
    
    With Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "vrm"
        .DialogTitle = "Create New GB VRAM - Choose Filename"
        .Filename = ""
        .Filter = "GB VRAMs (*.vrm)|*.vrm"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Set project path
        Dim i As Integer
        Dim flag As Boolean
        Dim sParentPath As String
        
        flag = False
        For i = Len(.Filename) To 1 Step -1
            If Mid$(.Filename, i, 1) = "\" Then
                If flag Then
                    sParentPath = Mid$(.Filename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Create new editor
        Dim frm As New frmEditVRAM
        Dim res As New clsGBVRAM
        
        gResourceCache.AddResourceToCache .Filename, res, frm
        Set frm.GBVRAM = res
        frm.bChanged = True
        frm.sFilename = .Filename
        frm.GBVRAM.intResource_ParentPath = sParentPath
        frm.Show
         
    'Organize forms
        CleanUpForms Me
        
    End With
    
    Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileNewVRAM_Click Error"

End Sub


Public Sub mnuFileOpen_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Open an unspecified file
'***************************************************************************
    
    Dim sFilename As String
    
    Dim dialogFrm As New frmDialog
    
    On Error Resume Next
    If gsProjectPath <> "" Then
        dialogFrm.InitDir = gsProjectPath
    Else
        dialogFrm.InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
    End If
    On Error GoTo HandleErrors
    
    dialogFrm.DialogTitle = "Open File"
    dialogFrm.Show vbModal
    If dialogFrm.bCancel Then
        Exit Sub
    End If
    
    Dim i As Integer
    
    For i = 1 To dialogFrm.FilenameCount
        
        gsCurPath = dialogFrm.Filenames(i)
        sFilename = dialogFrm.Filenames(i)
        
        If Not InStr(UCase$(sFilename), "C:\") <> 0 Then
            MsgBox "Opening files from the server is prohibited!  You must copy the appropriate files to your desktop and then open them again.", vbCritical, sFilename
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
            
    'Determine file type
        Dim nFilenum As Integer
        Dim nType As Byte
        nFilenum = FreeFile
        
        Open sFilename For Binary Access Read As #nFilenum
            Get #nFilenum, , nType
        Close #nFilenum
        
    'Set client variable to the appropriate type
        Dim client As Form
        
        Select Case nType
            Case GB_BITMAP
                Set client = New frmEditBitmap
            Case GB_MAP
                Set client = New frmEditMap
            Case GB_PALETTE
                Set client = New frmEditPalette
            Case GB_PATTERN
                Set client = New frmEditBackground
            Case GB_BG
                Set client = New frmEditBackground
            Case GB_VRAM
                Set client = New frmEditVRAM
            Case GB_COLLISIONMAP
                Set client = New frmEditCollisionMap
            Case GB_SPRITEGROUP
                Set client = New frmEditSpriteGroup
        End Select
        
        If client Is Nothing Then
            Screen.MousePointer = 0
            MsgBox "Error during load!  (File may be invalid)", vbCritical, "mdiMain:mnuFileOpen Error"
            Exit Sub
        End If
        
    'Set client's resource to the loaded resource
        Dim Resource As Object
        
        Set Resource = gResourceCache.GetResourceFromFile(sFilename, client)
        
        If Resource Is Nothing Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Select Case nType
            Case GB_BITMAP
                Set client.GBBitmap = Resource
            Case GB_MAP
                Set client.GBMap = Resource
                client.bOpening = True
            Case GB_PALETTE
                Set client.GBPalette = Resource
                client.bOpening = True
            Case GB_PATTERN
                Set client.GBBackground = Resource
                client.bOpening = True
                client.GBBackground.tBackgroundType = GB_PATTERNBG
            Case GB_BG
                Set client.GBBackground = Resource
                client.bOpening = True
                client.GBBackground.tBackgroundType = GB_RAWBG
            Case GB_VRAM
                Set client.GBVRAM = Resource
            Case GB_COLLISIONMAP
                Set client.GBCollisionMap = Resource
                Set client.GBCollisionMap.GBCollisionCodes = gGBCollisionCodes
                client.bOpening = True
            Case GB_SPRITEGROUP
                Set client.GBSpriteGroup = Resource
        End Select
        
    'Display editor
        
        client.sFilename = sFilename
        client.Show
        
    Next i
            
    Unload dialogFrm
    
    Screen.MousePointer = 0
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileOpenBatch_Click Error"
    Screen.MousePointer = 0
End Sub

Private Sub mnuFileOpenAdvancedBitmap_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Open an unspecified file
'***************************************************************************
    
    Dim sFilename As String
    
    With Dialog
    
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Open File"
        .Filename = ""
        .Filter = "GB Bitmaps (*.bit)|*.bit"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename

        sFilename = .Filename
        Screen.MousePointer = vbHourglass
            
    'Determine file type
        Dim nFilenum As Integer
        Dim nType As Byte
        nFilenum = FreeFile
        
        Open sFilename For Binary Access Read As #nFilenum
            Get #nFilenum, , nType
        Close #nFilenum
        
    'Set client variable to the appropriate type
        Dim client As Form
        
        Select Case nType
            Case GB_BITMAP
                Set client = New frmAdvBitmap
        End Select
        
        If client Is Nothing Then
            Screen.MousePointer = 0
            MsgBox "Error during load!  (File may be invalid)", vbCritical, "mdiMain:mnuFileOpen Error"
            Exit Sub
        End If
        
    'Set client's resource to the loaded resource
        Dim Resource As Object
        
        Set Resource = gResourceCache.GetResourceFromFile(.Filename, client)
        
        Select Case nType
            Case GB_BITMAP
                Set client.GBBitmap = Resource
        End Select
        
    'Display editor
        
        client.sFilename = sFilename
        client.Show
        
        Screen.MousePointer = 0
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileOpenAdvancedBitmap_Click Error"
    Screen.MousePointer = 0
End Sub


Private Sub mnuFileOpenNormal_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Open an unspecified file
'***************************************************************************
    
    Dim sFilename As String
    
    With Dialog

    'Get filename
        If gsProjectPath <> "" Then
            .InitDir = gsProjectPath
        Else
            .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        End If
        
        .DialogTitle = "Open File"
        .Filename = ""
        .Filter = "All Files (*.*)|*.*|GB Bitmaps (*.bit)|*.bit|GB Maps (*.map)|*.map|GB Palettes (*.pal)|*.pal|GB Patterns (*.pat)|*.pat|GB VRAMs (*.vrm)|*.vrm|GB BGs (*.bg)|*.bg|GB Collision Maps (*.clm)|*.clm|GB Sprite Groups (*.spr)|*.spr|GB Path Groups (*.pth)|*.pth"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
    
        sFilename = .Filename
        
        If Not InStr(UCase$(sFilename), "C:\") <> 0 Then
            MsgBox "Opening files from the server is prohibited!  You must copy the appropriate files to your desktop and then open them again.", vbCritical, sFilename
            Exit Sub
        End If
        
        If InStr(UCase$(sFilename), ".SOURCE") <> 0 Then
            MsgBox "You cannot open a .SOURCE file.  Please choose the original file!", vbInformation, "Load Error"
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
            
    'Determine file type
        Dim nFilenum As Integer
        Dim nType As Byte
        nFilenum = FreeFile
        
        Open sFilename For Binary Access Read As #nFilenum
            Get #nFilenum, , nType
        Close #nFilenum
        
    'Set client variable to the appropriate type
        Dim client As Form
        
        Select Case nType
            Case GB_BITMAP
                Set client = New frmEditBitmap
            Case GB_MAP
                Set client = New frmEditMap
            Case GB_PALETTE
                Set client = New frmEditPalette
            Case GB_PATTERN
                Set client = New frmEditBackground
            Case GB_BG
                Set client = New frmEditBackground
            Case GB_VRAM
                Set client = New frmEditVRAM
            Case GB_COLLISIONMAP
                Set client = New frmEditCollisionMap
            Case GB_SPRITEGROUP
                Set client = New frmEditSpriteGroup
        End Select
        
        If client Is Nothing Then
            Screen.MousePointer = 0
            MsgBox "Error during load!  (File may be invalid)", vbCritical, "mdiMain:mnuFileOpen Error"
            Exit Sub
        End If
        
    'Set client's resource to the loaded resource
        Dim Resource As Object
        
        Set Resource = gResourceCache.GetResourceFromFile(.Filename, client)
        
        If Resource Is Nothing Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Select Case nType
            Case GB_BITMAP
                Set client.GBBitmap = Resource
            Case GB_MAP
                Set client.GBMap = Resource
                client.bOpening = True
            Case GB_PALETTE
                Set client.GBPalette = Resource
                client.bOpening = True
            Case GB_PATTERN
                Set client.GBBackground = Resource
                client.bOpening = True
                client.GBBackground.tBackgroundType = GB_PATTERNBG
            Case GB_BG
                Set client.GBBackground = Resource
                client.bOpening = True
                client.GBBackground.tBackgroundType = GB_RAWBG
            Case GB_VRAM
                Set client.GBVRAM = Resource
            Case GB_COLLISIONMAP
                Set client.GBCollisionMap = Resource
                Set client.GBCollisionMap.GBCollisionCodes = gGBCollisionCodes
                client.bOpening = True
            Case GB_SPRITEGROUP
                Set client.GBSpriteGroup = Resource
        End Select
        
    'Display editor
        
        client.sFilename = sFilename
        client.Show
        
        Screen.MousePointer = 0
        
    End With
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuFileOpen_Click Error"
    Screen.MousePointer = 0
End Sub

Private Sub mnuGBOuputBankPacker_Click()

    frmBankPacker.Show

End Sub

Private Sub mnuHelpAbout_Click()

'***************************************************************************
'   Display the About window
'***************************************************************************
    
    frmAbout.Show
    
End Sub





Private Sub mnuNewProjectStructure_Click()

    Dim sPath As String
    sPath = GetPathDialog
    
    If sPath = "" Then
        Exit Sub
    End If
    
    Dim sProjectName As String
    sProjectName = InputBox("Enter a name for the project:", "Input Name")
    
    If sProjectName = "\" Then
        Exit Sub
    End If
    
    Dim str As String
    str = sPath & "\" & sProjectName
    
    CreateDir str
    CreateDir str & "\Bitmaps"
    CreateDir str & "\Collision"
    CreateDir str & "\Maps"
    CreateDir str & "\Palettes"
    CreateDir str & "\Patterns"
    CreateDir str & "\VRAMs"
    CreateDir str & "\Sprites"
    CreateDir str & "\Backgrounds"
    
    MsgBox "Project directories created successfully!", vbInformation, "Success"
    
End Sub


Private Sub mnuToolsAdvancedBitmapTool_Click()

    On Error GoTo HandleErrors

    Dim frm As New frmAdvBitmap
    frm.Show
        
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuToolsAdvancedBitmapTool_Click Error"
End Sub

Private Sub mnuToolsOutputMapPictureOutput_Click()

    On Error GoTo HandleErrors
    
    frmMapOutput.Show vbModal
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:mnuToolsOutputMapPictureOutput Error"
End Sub

Private Sub mnuToolsPackRLE_Click()

    frmPackRLE.Show vbModal

End Sub


Private Sub mnuViewCache_Click()

    mnuViewCache.Checked = Not mnuViewCache.Checked
    picResourceCacheView.Visible = mnuViewCache.Checked

End Sub

Public Sub tmrHotKeys_Timer()

    If KeyDown(vbKeyB) Then
    'Marquee Tool
        SelectTool GB_MARQUEE
        Exit Sub
    End If

    If KeyDown(vbKeyV) Then
    'Brush Tool
        SelectTool GB_BRUSH
        Exit Sub
    End If

    If KeyDown(vbKeyZ) Then
    'Zoom Tool
        SelectTool GB_ZOOM
        Exit Sub
    End If

    If KeyDown(vbKeyP) Then
    'Pointer Tool
        SelectTool GB_POINTER
        Exit Sub
    End If

    If KeyDown(vbKeyF) Then
    'Pointer Tool
        SelectTool GB_BUCKET
        Exit Sub
    End If

    If KeyDown(vbKeyR) Then
    'Replace Tool
        SelectTool GB_REPLACE
    End If

End Sub

Private Sub tmrUndo_Timer()

    On Error GoTo HandleErrors

    SelectTool GB_POINTER
    tmrUndo.Enabled = False

    'mnuEditUndo.Enabled = (gbUndoFlag And Not gUndoCache Is Nothing)
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "mdiMain:tmrUndo_Timer Error"
End Sub

Public Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    gnTool = Button.Index - 1

End Sub


