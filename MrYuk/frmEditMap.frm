VERSION 5.00
Begin VB.Form frmEditMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "frmEditMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   4935
      Begin VB.CommandButton cmdCrop 
         Caption         =   "&Crop"
         Height          =   360
         Left            =   2640
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   360
         Left            =   2640
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdBrowsePattern 
         Caption         =   "&Browse Pat..."
         Height          =   360
         Left            =   1320
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEditPattern 
         Caption         =   "&Edit Pat"
         Height          =   360
         Left            =   1320
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   360
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Save &As..."
         Height          =   360
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
      End
      Begin VB.Frame fraData 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3840
         TabIndex        =   5
         Top             =   0
         Width           =   1410
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   180
            Width           =   360
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
            TabIndex        =   10
            Top             =   0
            Width           =   360
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
            Index           =   3
            Left            =   60
            TabIndex        =   9
            Top             =   180
            Width           =   135
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
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   -15
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
            Index           =   1
            Left            =   600
            TabIndex        =   7
            Top             =   195
            Width           =   105
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
            Index           =   0
            Left            =   585
            TabIndex        =   6
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin VB.Timer tmrArrows 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3840
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   240
      LargeChange     =   16
      Left            =   0
      Max             =   63
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4320
      Width           =   4800
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   4335
      LargeChange     =   16
      Left            =   4800
      Max             =   63
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   0
      Width           =   4800
   End
   Begin VB.CommandButton cmdChangeSize 
      Height          =   240
      Left            =   4800
      Picture         =   "frmEditMap.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Toggle Size"
      Top             =   4320
      Width           =   255
   End
End
Attribute VB_Name = "frmEditMap"
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

    Private Const DEF_WIDTH = 5280 'In Twips
    Private Const DEF_HEIGHT = 5445 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mViewport As New clsViewport
    
    Private mbChanged As Boolean
    Private mbCached As Boolean
    Private mbOpening As Boolean
    Private msFilename As String
    Private mSelection As New clsSelection
    Private mnSize As Integer
    Private mbCropped As Boolean
    Private mbUndoing As Boolean
    Private mbActive As Boolean
    Private mbZoomOut As Boolean
    
'***************************************************************************
'   Resource object
'***************************************************************************

    Private mGBMap As New clsGBMap

Public Property Let bOpening(bNewValue As Boolean)

    mbOpening = bNewValue

End Property


Public Property Set GBMap(oNewValue As clsGBMap)

    Set mGBMap = oNewValue

End Property

Public Property Get GBMap() As clsGBMap

    Set GBMap = mGBMap

End Property

Private Function mFill(X As Integer, Y As Integer, value As Integer, startValue As Integer) As Boolean

    On Error GoTo HandleErrors
    
    mFill = True

    If X < 0 Or X > mGBMap.width - 1 Or Y < 0 Or Y > mGBMap.height - 1 Then
        Exit Function
    End If
    
    If gBufferMap.MapData(X + (Y * mGBMap.width) + 1) = value Then
        Exit Function
    End If
    
    If gBufferMap.MapData(X + (Y * mGBMap.width) + 1) = startValue Then
        gBufferMap.MapData(X + (Y * mGBMap.width) + 1) = value
    Else
        Exit Function
    End If
    
    Dim ret As Boolean
    
    ret = mFill(X + 1, Y, value, startValue)
    If ret = False Then
        mFill = False
        Exit Function
    End If
    ret = mFill(X - 1, Y, value, startValue)
    If ret = False Then
        mFill = False
        Exit Function
    End If
    ret = mFill(X, Y + 1, value, startValue)
    If ret = False Then
        mFill = False
        Exit Function
    End If
    ret = mFill(X, Y - 1, value, startValue)
    If ret = False Then
        mFill = False
        Exit Function
    End If
    
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:mFill Error"
    mFill = False
End Function

Private Function mFill2(X As Integer, Y As Integer, value As Integer, startValue As Integer) As Boolean

    On Error GoTo HandleErrors

    mFill2 = True

    If X < 0 Or X > mGBMap.width - 1 Or Y < 0 Or Y > mGBMap.height - 1 Then
        Exit Function
    End If
    
    If mGBMap.MapData(X + (Y * mGBMap.width) + 1) = value Then
        Exit Function
    End If
    
    If mGBMap.MapData(X + (Y * mGBMap.width) + 1) = startValue Then
        mGBMap.MapData(X + (Y * mGBMap.width) + 1) = value
    Else
        Exit Function
    End If
    
    Dim ret As Boolean
    
    ret = mFill2(X + 1, Y, value, startValue)
    If ret = False Then
        mFill2 = False
        Exit Function
    End If
    ret = mFill2(X - 1, Y, value, startValue)
    If ret = False Then
        mFill2 = False
        Exit Function
    End If
    ret = mFill2(X, Y + 1, value, startValue)
    If ret = False Then
        mFill2 = False
        Exit Function
    End If
    ret = mFill2(X, Y - 1, value, startValue)
    If ret = False Then
        mFill2 = False
        Exit Function
    End If
    
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:mFill2 Error"
    mFill2 = False
End Function

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Private Sub cmdBrowsePattern_Click()

    On Error GoTo HandleErrors
    
    With mdiMain.Dialog
        .InitDir = mGBMap.intResource_ParentPath & "\Patterns"
        .DialogTitle = "Load Pattern"
        .Filename = ""
        .Filter = "GB Patterns (*.pat)|*.pat"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        
        Dim sDir As String
        Dim sFilename As String
        
        sFilename = .Filename
        sDir = "Maps"
        
        gbOpeningChild = True
        Set mGBMap.GBBackground = gResourceCache.GetResourceFromFile(sFilename, mGBMap)
        gbOpeningChild = False
    
        mGBMap.sPatternFile = sFilename
    
    End With
    
    mbChanged = True
    intResourceClient_Update

Exit Sub

HandleErrors:
    If Err.Description = "Type mismatch" Then
        MsgBox "Mr. Yuk is trying to open the file " & sFilename & ".  This file appears to be invalid.  Please choose a replacement file.", vbCritical, "clsGBBackground:intResourceClient_Update Error"
    
        With mdiMain.Dialog
            .InitDir = mGBMap.intResource_ParentPath & "\" & sDir
            .DialogTitle = "Find Replacement"
            .Filename = ""
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            If .Filename = "" Then
                Exit Sub
            End If
            gsCurPath = .Filename
            
            If InStr(.Filename, mGBMap.intResource_ParentPath & "\" & sDir & "\") = 0 Then
                CopyFile .Filename, mGBMap.intResource_ParentPath & "\" & sDir & "\" & GetTruncFilename(.Filename), True
                MsgBox "The file was copied to the current project's " & sDir & " directory.", vbInformation, "File Moved"
            End If
            
            sFilename = .Filename
                        
            Resume
            
        End With
    
    Else

        MsgBox Err.Description, vbCritical, "frmEditMap:cmdBrowsePattern_Click Error"
    End If
End Sub

Private Sub cmdChangeSize_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Toggle between two sizes for the form
'***************************************************************************

    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
        Exit Sub
    End If

    If mnSize = 0 Then
        If mGBMap.width < (20 * 2) Or mGBMap.height < (18 * 2) Then
            picDisplay.width = 32 * mGBMap.width / 2
            picDisplay.height = 32 * mGBMap.height / 2
            mViewport.DisplayWidth = 32 * mGBMap.width / 2
            mViewport.DisplayHeight = 32 * mGBMap.height / 2
            mViewport.ViewportWidth = 16 * mGBMap.width
            mViewport.ViewportHeight = 16 * mGBMap.height
            hsbScroll.Max = mViewport.hScrollMax
            vsbScroll.Max = mViewport.vScrollMax
            Form_Resize
            intResourceClient_Update
            mnSize = 1
            Exit Sub
        End If
    End If
    
    If mnSize = 1 Then
        If mGBMap.width < 30 * 2 Or mGBMap.height < 27 * 2 Then
            mnSize = 2
        End If
    End If

    If mnSize = 0 Then
        picDisplay.width = 640
        picDisplay.height = 576
        mViewport.DisplayWidth = 640
        mViewport.DisplayHeight = 576
        mViewport.ViewportWidth = 320 * 2
        mViewport.ViewportHeight = 288 * 2
        mnSize = 1
    ElseIf mnSize = 1 Then
        picDisplay.width = 960
        picDisplay.height = 864
        mViewport.DisplayWidth = 960
        mViewport.DisplayHeight = 864
        mViewport.ViewportWidth = 480 * 2
        mViewport.ViewportHeight = 432 * 2
        mnSize = 2
    ElseIf mnSize = 2 Then
        picDisplay.width = 320
        picDisplay.height = 288
        mViewport.DisplayWidth = 320
        mViewport.DisplayHeight = 288
        mViewport.ViewportWidth = 160 * 2
        mViewport.ViewportHeight = 144 * 2
        mnSize = 0
    End If

    hsbScroll.Max = mViewport.hScrollMax
    vsbScroll.Max = mViewport.vScrollMax

    Form_Resize

    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:cmdChangeSize_Click Error"
End Sub

Private Sub cmdCrop_Click()

    Dim frm As New frmMapCrop
    
    Set frm.GBMap = mGBMap
    frm.Show vbModal
    
    If frm.bCancel Then
        Exit Sub
    End If
    
    lblW.Caption = Format$(mGBMap.width, "000")
    lblH.Caption = Format$(mGBMap.height, "000")
    
    With mViewport
        .Zoom = 1
        .DisplayWidth = 320 / 2
        .DisplayHeight = 288 / 2
        .ViewportWidth = 160 / 2
        .ViewportHeight = 144 / 2
        .SourceWidth = mGBMap.width * 16
        .SourceHeight = mGBMap.height * 16
        hsbScroll.Max = .hScrollMax
        vsbScroll.Max = .vScrollMax
    End With
    
    mnSize = 2
    cmdChangeSize_Click
    
    intResourceClient_Update

    mbCropped = True
    mbChanged = True

End Sub

Private Sub cmdEditPattern_Click()

'***************************************************************************
'   Open the pattern file in an editor
'***************************************************************************

    On Error GoTo HandleErrors
        
    Screen.MousePointer = vbHourglass
        
    Dim frm As New frmEditBackground
    
    frm.bOpening = True
    gResourceCache.AddResourceToCache mGBMap.sPatternFile, mGBMap.GBBackground, frm
    Set frm.GBBackground = mGBMap.GBBackground
    frm.sFilename = mGBMap.intResource_ParentPath & "\Patterns\" & GetTruncFilename(mGBMap.sPatternFile)
    frm.Show

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox "Load Error! (File format may be invalid)", vbCritical, "Load file error"
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current map into a .map file
'***************************************************************************

    PackFile mGBMap.intResource_ParentPath & "\Maps\" & GetTruncFilename(msFilename), mGBMap
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:cmdSave_Click Error"
End Sub


Private Sub cmdSaveAs_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current map under a new filename
'***************************************************************************

    Screen.MousePointer = vbHourglass

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Map"
        .Filename = ""
        .Filter = "GB Maps (*.map)|*.map"
        .ShowSave
        If .Filename = "" Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBMap
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBMap.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBMap, Me
        
        Dim frm As New frmEditBackground
        gbRefOpen = True
        frm.GBBackground.tBackgroundType = GB_PATTERNBG
        frm.bOpening = True
        Set frm.GBBackground = mGBMap.GBBackground
        Load frm
        Unload frm
        gbRefOpen = False
            
        mGBMap.Offscreen.Create mGBMap.width * 16, mGBMap.height * 16
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With
    
rExit:
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:cmdSaveAs_Click Error"
    GoTo rExit
End Sub



Private Sub cmdUndo_Click()

    On Error GoTo HandleErrors

    If mbUndoing Then
        Exit Sub
    End If

    mbUndoing = True

    If mbCropped Then
        mGBMap.SaveUndoState
        mbCropped = False
        Exit Sub
    End If

    mGBMap.RestoreUndoState
    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:cmdUndo_Click Error"
End Sub



Private Sub Form_Activate()

    mbActive = True

End Sub

Private Sub Form_Deactivate()

    mbActive = False

End Sub


Private Sub Form_Load()

    On Error GoTo HandleErrors

'***************************************************************************
'   Initialize the map editor
'***************************************************************************
    
'Set form dimensions
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
    If Not mbOpening Then
    'Display map properties
        frmMapProperties.nWidth = 64
        frmMapProperties.nHeight = 64
        frmMapProperties.sFilename = "(None)"
        frmMapProperties.ParentPath = mGBMap.intResource_ParentPath
        frmMapProperties.Show vbModal
    
        Screen.MousePointer = vbHourglass
        
        If frmMapProperties.bCancel Then
            gbClosingApp = True
            Unload frmMapProperties
            Unload Me
            gbClosingApp = False
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        mGBMap.width = frmMapProperties.nWidth
        mGBMap.height = frmMapProperties.nHeight
        mGBMap.sPatternFile = frmMapProperties.sFilename
    
        gbClosingApp = True
        Unload frmMapProperties
        gbClosingApp = False
    
        Screen.MousePointer = 0
        Set mGBMap.GBBackground = gResourceCache.GetResourceFromFile(mGBMap.sPatternFile, mGBMap)
        mGBMap.GBBackground.tBackgroundType = GB_PATTERNBG
    End If
    
    lblW.Caption = Format$(mGBMap.width, "000")
    lblH.Caption = Format$(mGBMap.height, "000")
    
    Dim frm As New frmEditBackground
    gbRefOpen = True
    frm.GBBackground.tBackgroundType = GB_PATTERNBG
    frm.bOpening = True
    Set frm.GBBackground = mGBMap.GBBackground
    Load frm
    Unload frm
    gbRefOpen = False
        
    mGBMap.Offscreen.Create mGBMap.width * 16, mGBMap.height * 16
        
    With mViewport
        .Zoom = 2
        
        .DisplayWidth = 320
        .DisplayHeight = 288
        
        .ViewportWidth = 160 * 2
        .ViewportHeight = 144 * 2
        
        .SourceWidth = mGBMap.width * 16
        .SourceHeight = mGBMap.height * 16
        
        hsbScroll.Max = .hScrollMax
        vsbScroll.Max = .vScrollMax
        
    End With

    intResourceClient_Update
    
    CleanUpForms Me

    mGBMap.UndoFlag = True
    mGBMap.SaveUndoState

    cmdChangeSize_Click
    gnTool = GB_ZOOM
    picDisplay_MouseDown vbRightButton, 0, 1, 1
    gnTool = GB_POINTER

    tmrArrows.Enabled = True

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:Form_Load Error"
End Sub

Private Sub mDrawGridLinesOnBank()

    On Error GoTo HandleErrors

'***************************************************************************
'   Create grid on display window
'***************************************************************************

    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim xCount As Integer
    Dim yCount As Integer
    
    For Y = 1 To mGBMap.Offscreen.height Step 16
        xCount = 0
        For X = 1 To mGBMap.Offscreen.width Step 16
        
        'Draw the large rectangles with bright color
            mGBMap.Offscreen.RECT X, Y, X + 16, Y + 16, vbBlack 'RGB(196, 141, 180)
            mGBMap.Offscreen.RECT X + 1, Y + 1, X + 15, Y + 15, vbWhite 'RGB(113, 64, 99)
        
            xCount = xCount + 1
        
        Next X
        yCount = yCount + 1
    Next Y

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:mDrawGridLinesOnBank Error"

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

Private Sub Form_Resize()

    On Error GoTo HandleErrors
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    vsbScroll.Left = picDisplay.width
    vsbScroll.height = picDisplay.height
    
    hsbScroll.Top = picDisplay.height
    hsbScroll.width = picDisplay.width
    
    cmdChangeSize.Top = vsbScroll.height
    cmdChangeSize.Left = hsbScroll.width
    
    fraButtons.Top = hsbScroll.Top + hsbScroll.height + 6
    
    Me.width = (vsbScroll.Left + vsbScroll.width + 6) * Screen.TwipsPerPixelX
    Me.height = (fraButtons.Top + fraButtons.height + 6 + 24) * Screen.TwipsPerPixelY
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:Form_Resize Error"
End Sub


Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo HandleErrors
    
    gResourceCache.ReleaseClient Me
 
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:Form_Unload Error"
End Sub
Private Sub hsbScroll_Change()

    On Error GoTo HandleErrors
    
    mViewport.ViewportX = hsbScroll.value
    
    picDisplay.Cls

    mViewport.Draw picDisplay.hdc, mGBMap.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:hsbScroll_Change Error"

End Sub

Private Sub hsbScroll_GotFocus()

    Me.SetFocus

End Sub


Private Sub hsbScroll_Scroll()

    hsbScroll_Change

End Sub


Public Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim patX As Integer
    Dim patY As Integer
    Dim val As Integer
    
    picDisplay.Cls
    mGBMap.Offscreen.Cls
    
    For Y = 0 To mGBMap.height - 1
        For X = 0 To mGBMap.width - 1
            val = mGBMap.MapData(X + (Y * mGBMap.width) + 1) + 1
            patX = ((val - 1) Mod 16) + 1
            patY = ((val - 1) \ 16) + 1
            If patX >= 0 And patY >= 0 Then
                mGBMap.GBBackground.MapOffscreen.BlitRaster mGBMap.Offscreen.hdc, X * 16, Y * 16, 16, 16, (patX - 1) * 16, (patY - 1) * 16, vbSrcCopy
            End If
        Next X
    Next Y
    
'If using the marquee tool, show the selected area inversed
    If gSelection.SrcForm Is Me Then
        mGBMap.Offscreen.BlitRaster mGBMap.Offscreen.hdc, gSelection.Left * gSelection.CellWidth, gSelection.Top * gSelection.CellHeight, gSelection.SelectionWidth * gSelection.CellWidth, gSelection.SelectionHeight * gSelection.CellHeight, 0, 0, vbDstInvert
    End If
    
    mViewport.Draw picDisplay.hdc, mGBMap.Offscreen
    
    Me.Caption = UCase$(GetTruncFilename(msFilename))
        
    picDisplay.Refresh
    
    If tType = GB_ACTIVEEDITOR Then
        If Not mGBMap Is Nothing Then
            mGBMap.UpdateClients Me
            mGBMap.UpdateClients Me
        End If
    End If

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:intResourceClient_Update Error"
End Sub



Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)

    Select Case gnTool
    
        Case GB_POINTER
    
        Case GB_MARQUEE
            
        'Create a selection where the user clicked
            mSelection.Left = click.X
            mSelection.Top = click.Y
            mSelection.Right = mSelection.Left + 1
            mSelection.Bottom = mSelection.Top + 1
            mSelection.AreaWidth = mGBMap.width
            mSelection.AreaHeight = mGBMap.height
            mSelection.CellWidth = 16
            mSelection.CellHeight = 16
            
            Set gSelection = mSelection.FixRect
            Set gSelection.SrcForm = Me
            
            intResourceClient_Update
    
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
        
            SelectTool GB_POINTER
    
        Case GB_BRUSH, GB_BUCKET
    
            If gSelection.SrcForm Is Nothing Then
                Exit Sub
            End If
            
            Dim ret As Integer
            Dim retW As Integer
            Dim retX As Integer
            Dim retY As Integer
            Dim xIndex As Integer
            Dim yIndex As Integer
            
            If TypeOf gSelection.SrcForm Is frmEditMap Then
                
                mGBMap.SaveUndoState
                mGBMap.UndoFlag = False
                
                ret = gSelection.GetFirstElement
            
                Do Until ret < 0
                    xIndex = click.X + gSelection.CursorX
                    yIndex = click.Y + gSelection.CursorY
                    If xIndex < 0 Or yIndex < 0 Or xIndex >= mGBMap.width Or yIndex >= mGBMap.height Then
                    Else
                        retX = ((ret - 1) Mod mGBMap.width)
                        retY = ((ret - 1) \ mGBMap.width)
                        If gnTool = GB_BUCKET Then
                            mFill2 click.X, click.Y, gBufferMap.MapData(retX + (retY * (mGBMap.width)) + 1), GBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1)
                            'mFill click.X, click.Y, gBufferMap.MapData(retX + (retY * (mGBMap.width)) + 1), mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1)
                            intResourceClient_Update
                        Else
                            mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1) = gBufferMap.MapData(retX + (retY * (mGBMap.width)) + 1)
                            
                            Dim val As Integer
                            Dim patX As Integer
                            Dim patY As Integer
                            Dim dw As Integer
                            Dim dh As Integer
                            val = mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1) + 1
                            dw = mViewport.DisplayWidth \ mViewport.ViewportWidth
                            dh = mViewport.DisplayHeight \ mViewport.ViewportHeight
                            patX = ((val - 1) Mod 16) + 1
                            patY = ((val - 1) \ 16) + 1
                            If patX >= 0 And patY >= 0 Then
                                If gSelection.SelectionWidth = 1 And gSelection.SelectionHeight = 1 Then
                                    mGBMap.GBBackground.MapOffscreen.BlitZoom picDisplay.hdc, ((xIndex * 16 * dw) - (mViewport.ViewportX * dw)) * mViewport.Zoom, ((yIndex * 16 * dh) - (mViewport.ViewportY * dh)) * mViewport.Zoom, 16 * dw * mViewport.Zoom, 16 * dh * mViewport.Zoom, (patX - 1) * 16, (patY - 1) * 16, 16, 16
                                    mGBMap.GBBackground.MapOffscreen.BlitRaster mGBMap.Offscreen.hdc, xIndex * 16, yIndex * 16, 16, 16, (patX - 1) * 16, (patY - 1) * 16, vbSrcCopy
                                End If
                            End If
                            picDisplay.Refresh
                            
                        End If
                    End If
                    ret = gSelection.GetNextElement
                Loop
                
                If gSelection.SelectionWidth > 1 Or gSelection.SelectionHeight > 1 Then
                    intResourceClient_Update
                End If
                
            Else
            
                If Not TypeOf gSelection.SrcForm Is frmEditBackground Then
                    Exit Sub
                End If
                
            'Transfer data from pattern to map
                
                mGBMap.SaveUndoState
                mGBMap.UndoFlag = False
                    
                ret = gSelection.GetFirstElement
               
                Do Until ret < 0
                    xIndex = click.X + gSelection.CursorX
                    yIndex = click.Y + gSelection.CursorY
                    If xIndex < 0 Or yIndex < 0 Or xIndex >= mGBMap.width Or yIndex >= mGBMap.height Then
                    Else
                        If gnTool = GB_BUCKET Then
                            mFill2 click.X, click.Y, ret - 1, mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1)
                            intResourceClient_Update
                        Else
                            mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1) = ret - 1
                            
                            val = mGBMap.MapData(xIndex + (yIndex * (mGBMap.width)) + 1) + 1
                            dw = mViewport.DisplayWidth \ mViewport.ViewportWidth
                            dh = mViewport.DisplayHeight \ mViewport.ViewportHeight
                            patX = ((val - 1) Mod 16) + 1
                            patY = ((val - 1) \ 16) + 1
                            If patX >= 0 And patY >= 0 Then
                                If gSelection.SelectionWidth = 1 And gSelection.SelectionHeight = 1 Then
                                    mGBMap.GBBackground.MapOffscreen.BlitZoom picDisplay.hdc, ((xIndex * 16 * dw) - (mViewport.ViewportX * dw)) * mViewport.Zoom, ((yIndex * 16 * dh) - (mViewport.ViewportY * dh)) * mViewport.Zoom, 16 * dw * mViewport.Zoom, 16 * dh * mViewport.Zoom, (patX - 1) * 16, (patY - 1) * 16, 16, 16
                                    mGBMap.GBBackground.MapOffscreen.BlitRaster mGBMap.Offscreen.hdc, xIndex * 16, yIndex * 16, 16, 16, (patX - 1) * 16, (patY - 1) * 16, vbSrcCopy
                                End If
                            End If
                            picDisplay.Refresh
                            
                        End If
                    End If
                    ret = gSelection.GetNextElement
                Loop
            
                If gSelection.SelectionWidth > 1 Or gSelection.SelectionHeight > 1 Then
                    intResourceClient_Update
                End If
            
            End If
        
        'Set flag used for saving
            mbChanged = True
            
        Case GB_REPLACE
        
            Dim i As Integer
            Dim dummy As Integer
            
            dummy = mGBMap.MapData(click.X + (click.Y * mGBMap.width) + 1)
                        
            For i = 0 To (mGBMap.width * mGBMap.height)
                If mGBMap.MapData(i) = dummy Then
                    mGBMap.MapData(i) = gnReplaceTile
                End If
            Next i
            
            mbChanged = True
            
            intResourceClient_Update
        
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:picDisplay_MouseDown Error"
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

    If X >= picDisplay.width Or X < 0 Or Y >= picDisplay.height Or Y < 0 Then
        Exit Sub
    End If

    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)
    
    lblX.Caption = Format$(CStr(click.X), "000")
    lblY.Caption = Format$(CStr(click.Y), "000")

    If Button = vbLeftButton And gnTool <> GB_ZOOM And gnTool <> GB_MARQUEE Then
        picDisplay_MouseDown Button, Shift, X, Y
    End If

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER

            picDisplay.MousePointer = vbArrow
            
        Case GB_MARQUEE
        
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
                
                intResourceClient_Update
            
            End If
            
        Case GB_BRUSH
            
            picDisplay.MouseIcon = mdiMain.picDragAdd.Picture
            picDisplay.MousePointer = vbCustom
    
        Case GB_ZOOM
    
            picDisplay.MouseIcon = mdiMain.picZoom.Picture
            picDisplay.MousePointer = vbCustom
    
        Case GB_SETTER
        
            picDisplay.MousePointer = vbArrow
            
        Case GB_BUCKET
    
            picDisplay.MousePointer = vbUpArrow
            
        Case GB_REPLACE
        
            picDisplay.MousePointer = vbUpArrow
    
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:picDisplay_MouseMove Error"
End Sub


Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
            
        Case GB_MARQUEE
        
            Dim i As Integer
            Dim dx As Integer
            Dim dy As Integer
        
            gBufferMap.width = gSelection.SrcForm.GBMap.width
            gBufferMap.height = gSelection.SrcForm.GBMap.height
        
            For dy = 1 To gSelection.SrcForm.GBMap.height
                For dx = 1 To gSelection.SrcForm.GBMap.width
                    i = dx + ((dy - 1) * gSelection.SrcForm.GBMap.width)
                    gBufferMap.MapData(i) = gSelection.SrcForm.GBMap.MapData(i)
                Next dx
            Next dy
            
            SelectTool GB_BRUSH
                
        Case GB_BRUSH
        
            mGBMap.UndoFlag = True
            SelectTool GB_BRUSH
            
        Case GB_BUCKET
        
            mGBMap.UndoFlag = True
            SelectTool GB_BUCKET
        
        Case GB_ZOOM
        
    End Select

End Sub



Private Sub tmrArrows_Timer()

    If mbActive = False Then
        Exit Sub
    End If

    If hsbScroll.value - 8 >= 0 Then
        If KeyDown(vbKeyLeft) Then
            hsbScroll.value = hsbScroll.value - 8
        End If
    Else
        If KeyDown(vbKeyLeft) Then
            hsbScroll.value = 0
        End If
    End If

    If hsbScroll.value + 8 <= hsbScroll.Max Then
        If KeyDown(vbKeyRight) Then
            hsbScroll.value = hsbScroll.value + 8
        End If
    Else
        If KeyDown(vbKeyRight) Then
            hsbScroll.value = hsbScroll.Max
        End If
    End If

    If vsbScroll.value - 8 >= 0 Then
        If KeyDown(vbKeyUp) Then
            vsbScroll.value = vsbScroll.value - 8
        End If
    Else
        If KeyDown(vbKeyUp) Then
            vsbScroll.value = 0
        End If
    End If

    If vsbScroll.value + 8 <= vsbScroll.Max Then
        If KeyDown(vbKeyDown) Then
            vsbScroll.value = vsbScroll.value + 8
        End If
    Else
        If KeyDown(vbKeyDown) Then
            vsbScroll.value = vsbScroll.Max
        End If
    End If
    
    If KeyDown(vbKeyF5) Then
        intResourceClient_Update
    End If

    If KeyDown(vbKeyU) Then
        cmdUndo_Click
    Else
        mbUndoing = False
    End If

End Sub

Private Sub vsbScroll_Change()

    On Error GoTo HandleErrors

    mViewport.ViewportY = vsbScroll.value
    
    picDisplay.Cls
    
    mViewport.Draw picDisplay.hdc, mGBMap.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:vsbScroll_Change Error"
End Sub


Private Sub vsbScroll_GotFocus()

    Me.SetFocus

End Sub


Private Sub vsbScroll_Scroll()

    vsbScroll_Change

End Sub


