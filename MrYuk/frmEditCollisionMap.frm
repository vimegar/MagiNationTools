VERSION 5.00
Begin VB.Form frmEditCollisionMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collision Map"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmEditCollisionMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   5055
      Begin VB.CheckBox chkReplaceToMap 
         Caption         =   "&Replace to Map"
         Height          =   345
         Left            =   3555
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   495
         Width           =   1455
      End
      Begin VB.CheckBox chkViewCodes 
         Caption         =   "&View Codes"
         Height          =   225
         Left            =   3555
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton cmdColRect 
         Caption         =   "&Output HS..."
         Height          =   360
         Left            =   2280
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   1140
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   360
         Left            =   1170
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   360
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   960
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Save &As..."
         Height          =   360
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdAddBorder 
         Caption         =   "&Add Border"
         Height          =   360
         Left            =   1170
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.CheckBox chkFillType 
         Caption         =   "&Fill to Map"
         Height          =   345
         Left            =   3555
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   210
         Width           =   1095
      End
      Begin VB.Frame fraData 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2160
         TabIndex        =   5
         Top             =   0
         Width           =   1290
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
            Left            =   675
            TabIndex        =   13
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
            Index           =   1
            Left            =   690
            TabIndex        =   12
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
            Index           =   2
            Left            =   60
            TabIndex        =   11
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
            Index           =   3
            Left            =   60
            TabIndex        =   10
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
            Left            =   285
            TabIndex        =   9
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
            Left            =   285
            TabIndex        =   8
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
            Left            =   915
            TabIndex        =   7
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
            Left            =   915
            TabIndex        =   6
            Top             =   180
            Width           =   360
         End
      End
   End
   Begin VB.Timer tmrArrows 
      Interval        =   1
      Left            =   3720
      Top             =   3480
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
      Picture         =   "frmEditCollisionMap.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Toggle Size"
      Top             =   4320
      Width           =   255
   End
End
Attribute VB_Name = "frmEditCollisionMap"
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

    Private Const DEF_WIDTH = 5235 'In Twips
    Private Const DEF_HEIGHT = 5505 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mViewport As New clsViewport
    
    Private mbChanged As Boolean
    Private mbCached As Boolean
    Private mbOpening As Boolean
    Private msFilename As String
    Private mnSize As Integer
    Private mbUndoing As Boolean
    Private mbActive As Boolean
    
'***************************************************************************
'   Resource object
'***************************************************************************

    Private mGBCollisionMap As New clsGBCollisionMap
Private Sub mFill(X As Integer, Y As Integer, value As Integer, startValue As Integer)

    On Error GoTo HandleErrors

    If X < 0 Or X > mGBCollisionMap.GBMap.width - 1 Or Y < 0 Or Y > mGBCollisionMap.GBMap.height - 1 Then
        Exit Sub
    End If
    
    If mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = value Then
        Exit Sub
    End If
    
    If mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = startValue Then
        mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = value
    Else
        Exit Sub
    End If
    
    mFill X + 1, Y, value, startValue
    mFill X - 1, Y, value, startValue
    mFill X, Y + 1, value, startValue
    mFill X, Y - 1, value, startValue
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:mFill Error"
End Sub

Public Property Let bOpening(bNewValue As Boolean)

    mbOpening = bNewValue

End Property


Public Property Set GBCollisionMap(oNewValue As clsGBCollisionMap)

    Set mGBCollisionMap = oNewValue

End Property

Public Property Get GBCollisionMap() As clsGBCollisionMap

    Set GBCollisionMap = mGBCollisionMap

End Property

Private Sub mFillToMap(X As Integer, Y As Integer, value As Integer, startValue As Integer)
    
    On Error GoTo HandleErrors

    If X < 0 Or X > mGBCollisionMap.GBMap.width - 1 Or Y < 0 Or Y > mGBCollisionMap.GBMap.height - 1 Then
        Exit Sub
    End If
    
    If mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = value Then
        Exit Sub
    End If
    
    If mGBCollisionMap.GBMap.MapData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = startValue Then
        mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1) = value
    Else
        Exit Sub
    End If
    
    mFillToMap X + 1, Y, value, startValue
    mFillToMap X - 1, Y, value, startValue
    mFillToMap X, Y + 1, value, startValue
    mFillToMap X, Y - 1, value, startValue
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditMap:mFill Error"
End Sub

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Private Sub chkViewCodes_Click()

    intResourceClient_Update
    
End Sub

Private Sub cmdAddBorder_Click()

    Dim X As Integer
    Dim Y As Integer
    Dim w As Integer
    Dim h As Integer
    
    w = mGBCollisionMap.GBMap.width
    h = mGBCollisionMap.GBMap.height
    
    For Y = 0 To (h - 1)
        For X = 0 To (w - 1)
            If X = 0 Or X = (w - 1) Or Y = 0 Or Y = (h - 1) Then
                mGBCollisionMap.CollisionData(X + (Y * w) + 1) = 8
            End If
        Next X
    Next Y

    intResourceClient_Update

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
        If mGBCollisionMap.GBMap.width < (20 * 2) Or mGBCollisionMap.GBMap.height < (18 * 2) Then
            picDisplay.width = 32 * mGBCollisionMap.GBMap.width / 2
            picDisplay.height = 32 * mGBCollisionMap.GBMap.height / 2
            mViewport.DisplayWidth = 32 * mGBCollisionMap.GBMap.width / 2
            mViewport.DisplayHeight = 32 * mGBCollisionMap.GBMap.height / 2
            mViewport.ViewportWidth = 16 * mGBCollisionMap.GBMap.width
            mViewport.ViewportHeight = 16 * mGBCollisionMap.GBMap.height
            hsbScroll.Max = mViewport.hScrollMax
            vsbScroll.Max = mViewport.vScrollMax
            Form_Resize
            intResourceClient_Update
            mnSize = 1
            Exit Sub
        End If
    End If
    
    If mnSize = 1 Then
        If mGBCollisionMap.GBMap.width < 30 * 2 Or mGBCollisionMap.GBMap.height < 27 * 2 Then
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
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:cmdChangeSize_Click Error"
End Sub




Private Sub cmdColRect_Click()

    On Error GoTo HandleErrors

    Dim sFilename As String

    'With mdiMain.Dialog
    '    .DefaultExt = "mgi"
    '    .DialogTitle = "Save As MGI File"
    '    .filename = ""
    '    .Filter = "Magi Parse File (*.mgi)|*.mgi"
    '    .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
    '    .ShowSave
    '    If .filename = "" Then
    '        Exit Sub
    '    End If
    '    gsCurPath = .filename
    '
    '    Dim nFilenum As Integer
    '    nFilenum = FreeFile
    '
    '    Open .filename For Output As #nFilenum
    '    sFilename = .filename
    '
    'End With

    sFilename = App.Path & "\hsrect.mgi"

    Dim nFilenum As Integer
    nFilenum = FreeFile

    Open sFilename For Output As #nFilenum

    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Index As Integer
    Dim bMatched As Boolean

    Screen.MousePointer = vbHourglass

    ReDim colRects(0) As New clsHotspotRect
    
    For Y = 0 To (mGBCollisionMap.GBMap.height - 1)
        For X = 0 To (mGBCollisionMap.GBMap.width - 1)

            Index = X + (Y * mGBCollisionMap.GBMap.width) + 1
            
            If mGBCollisionMap.CollisionData(Index) >= 191 Then
            
                bMatched = False
                
                For i = 1 To UBound(colRects)
                    If mGBCollisionMap.CollisionData(Index) = colRects(i).nType Then
                        bMatched = True
                        If X < colRects(i).nLeft Then colRects(i).nLeft = X
                        If X > colRects(i).nRight Then colRects(i).nRight = X
                        If Y < colRects(i).nTop Then colRects(i).nTop = Y
                        If Y > colRects(i).nBottom Then colRects(i).nBottom = Y
                        Exit For
                    End If
                Next i
                
                If Not bMatched Then
                    ReDim Preserve colRects(UBound(colRects) + 1)
                    colRects(UBound(colRects)).nType = mGBCollisionMap.CollisionData(Index)
                    colRects(UBound(colRects)).nLeft = X
                    colRects(UBound(colRects)).nRight = X
                    colRects(UBound(colRects)).nTop = Y
                    colRects(UBound(colRects)).nBottom = Y
                End If
                
            End If

        Next X
    Next Y

    QuicksortHSRect colRects, 1, UBound(colRects)

    'output to mgi file
    Dim mapName As String
    
    For i = 1 To Len(msFilename)
        If Mid$(msFilename, i, 1) = "_" Then
            mapName = Mid$(msFilename, i + 1)
            Exit For
        End If
    Next i
    
    For i = Len(mapName) To 1 Step -1
        If Mid$(mapName, i, 1) = "." Then
            mapName = Mid$(mapName, 1, i - 1)
            Exit For
        End If
    Next i
    
    For i = 1 To UBound(colRects)
        Print #nFilenum, "?_" & mapName & "_DR_" & CStr(i)
        Print #nFilenum, vbTab & vbTab & "HeroToDoor" & vbTab & "(" & CStr(colRects(i).nLeft) & "," & CStr(colRects(i).nTop) & ",0,0)"
        Print #nFilenum, ""
    Next i
    
    Close #nFilenum
    
    Shell "C:\Program Files\Metrowerks\CodeWarrior\bin\IDE.EXE " & sFilename
    
    DeleteFile sFilename
    
    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:cmdColRect_Click Error"
End Sub

Private Sub cmdFindReplace_Click()

    mGBCollisionMap.SaveUndoState

    Dim frm As New frmColFindReplace
    
    Set frm.GBCollisionMap = mGBCollisionMap
    frm.Show vbModal

    intResourceClient_Update

End Sub

Private Sub cmdSave_Click()

'***************************************************************************
'   Save the current collision map into a .clm file
'***************************************************************************

    On Error GoTo HandleErrors

    PackFile mGBCollisionMap.intResource_ParentPath & "\Collision\" & GetTruncFilename(msFilename), mGBCollisionMap
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:cmdSave_Click Error"
End Sub


Private Sub cmdSaveAs_Click()

'***************************************************************************
'   Save the current collision map under a new filename
'***************************************************************************

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Collision Map"
        .Filename = ""
        .Filter = "GB Collision Maps (*.clm)|*.clm"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBCollisionMap
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBCollisionMap.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBCollisionMap, Me
        
        Dim frm As New frmEditBackground
        gbRefOpen = True
        frm.GBBackground.tBackgroundType = GB_PATTERNBG
        frm.bOpening = True
        Set frm.GBBackground = mGBCollisionMap.GBMap.GBBackground
        Load frm
        Unload frm
        gbRefOpen = False
            
        mGBCollisionMap.GBMap.Offscreen.Create mGBCollisionMap.GBMap.width * 16, mGBCollisionMap.GBMap.height * 16
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:cmdSaveAs_Click Error"
End Sub


Private Sub cmdUndo_Click()

    If mbUndoing Then
        Exit Sub
    End If

    mbUndoing = True

    mGBCollisionMap.RestoreUndoState
    intResourceClient_Update

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
    
'Display map properties
    If Not mbOpening Then
        frmCollisionMapProperties.sFilename = "(None)"
        frmCollisionMapProperties.Show vbModal
    
        If frmCollisionMapProperties.bCancel Then
            gbClosingApp = True
            Unload frmCollisionMapProperties
            Unload Me
            gbClosingApp = False
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        mGBCollisionMap.sMapFile = frmCollisionMapProperties.sFilename
    
        gbClosingApp = True
        Unload frmCollisionMapProperties
        gbClosingApp = False
    End If
    
    Set mGBCollisionMap.GBMap = gResourceCache.GetResourceFromFile(mGBCollisionMap.sMapFile, mGBCollisionMap)
            
    If mGBCollisionMap.GBMap.sPatternFile <> "" Then
        Set mGBCollisionMap.GBMap.GBBackground = gResourceCache.GetResourceFromFile(mGBCollisionMap.GBMap.sPatternFile, mGBCollisionMap.GBMap)
        mGBCollisionMap.GBMap.GBBackground.tBackgroundType = GB_PATTERNBG
    End If
    
    Dim frm As New frmEditBackground
    gbRefOpen = True
    frm.GBBackground.tBackgroundType = GB_PATTERNBG
    frm.bOpening = True
    Set frm.GBBackground = mGBCollisionMap.GBMap.GBBackground
    Load frm
    Unload frm
    gbRefOpen = False
        
    mGBCollisionMap.GBMap.Offscreen.Create mGBCollisionMap.GBMap.width * 16, mGBCollisionMap.GBMap.height * 16
        
    With mViewport
        .Zoom = 2
        
        .DisplayWidth = 320
        .DisplayHeight = 288
        
        .SourceWidth = mGBCollisionMap.GBMap.width * 16
        .SourceHeight = mGBCollisionMap.GBMap.height * 16
        
        .ViewportWidth = 160 * 2
        .ViewportHeight = 144 * 2
        
        hsbScroll.Max = .hScrollMax
        vsbScroll.Max = .vScrollMax
        
    End With
    
    intResourceClient_Update
    
    CleanUpForms Me

    lblW.Caption = Format$(mGBCollisionMap.GBMap.width, "000")
    lblH.Caption = Format$(mGBCollisionMap.GBMap.height, "000")

    mGBCollisionMap.UndoFlag = True
    mGBCollisionMap.SaveUndoState
    
    cmdChangeSize_Click
    gnTool = GB_ZOOM
    picDisplay_MouseDown vbRightButton, 0, 1, 1
    gnTool = GB_POINTER

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:Form_Load Error"
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
    
    For Y = 1 To mGBCollisionMap.GBMap.Offscreen.height Step 16
        xCount = 0
        For X = 1 To mGBCollisionMap.GBMap.Offscreen.width Step 16
        
        'Draw the large rectangles with bright color
            mGBCollisionMap.GBMap.Offscreen.RECT X, Y, X + 16, Y + 16, vbBlack 'RGB(196, 141, 180)
            mGBCollisionMap.GBMap.Offscreen.RECT X + 1, Y + 1, X + 15, Y + 15, vbWhite 'RGB(113, 64, 99)
        
            xCount = xCount + 1
        
        Next X
        yCount = yCount + 1
    Next Y

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:mDrawGridLinesOnBank Error"
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
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:Form_Resize Error"
End Sub


Private Sub Form_Unload(Cancel As Integer)
 
    On Error GoTo HandleErrors
   
    'Dim i As Integer
    
    'For i = 0 To Forms.count - 1
    '    If TypeOf Forms(i) Is frmEditCollisionCodes Then
    '        If Forms(i).GBCollisionCodes Is mGBCollisionMap.GBCollisionCodes Then
    '            Unload Forms(i)
    '            Exit For
    '        End If
    '    End If
    'Next i
   
    
    gResourceCache.ReleaseClient Me

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:Form_Unload Error"
End Sub


Private Sub hsbScroll_Change()

    On Error GoTo HandleErrors

    mViewport.ViewportX = hsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBCollisionMap.GBMap.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:hsbScroll_Change Error"
End Sub

Private Sub hsbScroll_GotFocus()

    Me.SetFocus
    
End Sub


Private Sub hsbScroll_Scroll()

    hsbScroll_Change

End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim patX As Integer
    Dim patY As Integer
    Dim val As Integer
    
    picDisplay.Cls
    mGBCollisionMap.GBMap.Offscreen.Cls
    
    For Y = 0 To mGBCollisionMap.GBMap.height - 1
        For X = 0 To mGBCollisionMap.GBMap.width - 1
            val = mGBCollisionMap.GBMap.MapData(X + (Y * mGBCollisionMap.GBMap.width) + 1) + 1
            patX = (val - 1) Mod 16
            patY = (val - 1) \ 16
            If patX >= 0 And patY >= 0 Then
                mGBCollisionMap.GBMap.GBBackground.MapOffscreen.BlitRaster mGBCollisionMap.GBMap.Offscreen.hdc, X * 16, Y * 16, 16, 16, patX * 16, patY * 16, vbSrcCopy
            End If
            val = mGBCollisionMap.CollisionData(X + (Y * mGBCollisionMap.GBMap.width) + 1)
            patX = val Mod 16
            patY = val \ 16
            
            If chkViewCodes.value = vbChecked Then
                BitBlt mGBCollisionMap.GBMap.Offscreen.hdc, X * 16, Y * 16, 16, 16, gFrmColCodes.MaskHDC, patX * 16, patY * 16, vbSrcAnd
                BitBlt mGBCollisionMap.GBMap.Offscreen.hdc, X * 16, Y * 16, 16, 16, gFrmColCodes.PicHDC, patX * 16, patY * 16, vbSrcPaint
            End If
                    
        Next X
    Next Y
    
    mViewport.Draw picDisplay.hdc, mGBCollisionMap.GBMap.Offscreen
    
    Me.Caption = GetTruncFilename(msFilename)
        
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:intResourceClient_Update Error"
End Sub
Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

    If X >= picDisplay.width Or X < 0 Or Y >= picDisplay.height Or Y < 0 Then
        Exit Sub
    End If

    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
    
        Case GB_MARQUEE
    
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
            
            If Not TypeOf gSelection.SrcForm Is frmEditCollisionCodes Then
                Exit Sub
            End If
            
            mGBCollisionMap.SaveUndoState
            mGBCollisionMap.UndoFlag = False

        'Transfer data from codes to map
            If click.X < 0 Or click.Y < 0 Then
                Exit Sub
            End If
            
            If gnTool = GB_BUCKET Then
                If chkFillType.value = vbUnchecked Then
                    mFill click.X, click.Y, gSelection.SrcForm.SelectedCode, mGBCollisionMap.CollisionData(click.X + (click.Y * mGBCollisionMap.GBMap.width) + 1)
                Else
                    mFillToMap click.X, click.Y, gSelection.SrcForm.SelectedCode, mGBCollisionMap.GBMap.MapData(click.X + (click.Y * mGBCollisionMap.GBMap.width) + 1)
                End If
            Else
                mGBCollisionMap.CollisionData(click.X + (click.Y * mGBCollisionMap.GBMap.width) + 1) = gSelection.SrcForm.SelectedCode
            End If
            
        'Update visual
            intResourceClient_Update
        
        'Set flag used for saving
            mbChanged = True
    
        Case GB_REPLACE
        
            mGBCollisionMap.SaveUndoState
            mGBCollisionMap.UndoFlag = False
        
            If chkReplaceToMap.value = vbChecked Then
                
                Dim i As Integer
                Dim dummy As Integer
            
                dummy = mGBCollisionMap.GBMap.MapData(click.X + (click.Y * mGBCollisionMap.GBMap.width) + 1)
                            
                For i = 0 To (mGBCollisionMap.GBMap.width * mGBCollisionMap.GBMap.height)
                    If mGBCollisionMap.GBMap.MapData(i) = dummy Then
                        mGBCollisionMap.CollisionData(i) = gnReplaceTile
                    End If
                Next i
            
            Else
                
                dummy = mGBCollisionMap.CollisionData(click.X + (click.Y * mGBCollisionMap.GBMap.width) + 1)
                            
                For i = 0 To (mGBCollisionMap.GBMap.width * mGBCollisionMap.GBMap.height)
                    If mGBCollisionMap.CollisionData(i) = dummy Then
                        mGBCollisionMap.CollisionData(i) = gnReplaceTile
                    End If
                Next i
               
            End If
            
            mbChanged = True
            
            intResourceClient_Update
    
    End Select

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:picDisplay_MouseDown Error"
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

    If X >= picDisplay.width Or X <= 0 Or Y >= picDisplay.height Or Y <= 0 Then
        Exit Sub
    End If

    Dim click As New clsPoint
    Set click = mViewport.ScreenClickToGridClick(CInt(X), CInt(Y), 16, 16)

    lblX.Caption = Format$(CStr(click.X), "000")
    lblY.Caption = Format$(CStr(click.Y), "000")

    If Button = vbLeftButton And gnTool <> GB_ZOOM Then
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
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:picDisplay_MouseMove Error"
End Sub


Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
            
        Case GB_MARQUEE
        
        Case GB_BRUSH
        
            mGBCollisionMap.UndoFlag = True
            SelectTool GB_BRUSH
        
        Case GB_BUCKET
        
            mGBCollisionMap.UndoFlag = True
            SelectTool GB_BUCKET
        
        Case GB_ZOOM
        
        Case GB_REPLACE
        
            mGBCollisionMap.UndoFlag = True
        
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
    
    If KeyDown(vbKeyU) Then
        cmdUndo_Click
    Else
        mbUndoing = False
    End If

End Sub

Private Sub vsbScroll_Change()

    On Error GoTo HandleErrors

    mViewport.ViewportY = vsbScroll.value

    mViewport.Draw picDisplay.hdc, mGBCollisionMap.GBMap.Offscreen
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionMap:vsbScroll_Change Error"
End Sub


Private Sub vsbScroll_GotFocus()

    Me.SetFocus

End Sub


Private Sub vsbScroll_Scroll()

    vsbScroll_Change

End Sub


