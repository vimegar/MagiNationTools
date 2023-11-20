VERSION 5.00
Begin VB.Form frmEditBitmap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitmap"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmEditBitmap.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   Begin VB.PictureBox picWhite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   4320
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   120
      Width           =   4500
   End
   Begin VB.TextBox txtNumTiles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkNumTiles 
      Caption         =   "Define # of Tiles"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   0
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As..."
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1725
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1725
   End
   Begin VB.CommandButton cmdGetWindowsBitmap 
      Caption         =   "&Get BMP..."
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   3960
      Width           =   1725
   End
End
Attribute VB_Name = "frmEditBitmap"
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

    Private Const DEF_WIDTH = 3930 'In Twips
    Private Const DEF_HEIGHT = 5235 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mViewport As New clsViewport
    
    Private mbCached As Boolean
    Private mbChanged As Boolean
    Private msFilename As String
    Private msBMPfile As String
    Private moBMPOffscreen As New clsOffscreen
    
    Private miPalMap() As Byte
    
'***************************************************************************
'   Resource Object
'***************************************************************************

    Private mGBBitmap As clsGBBitmap
    Private mGBPalette As clsGBPalette

Public Property Get bChanged() As Boolean

    bChanged = mbChanged

End Property

Public Property Let bChanged(bNewValue As Boolean)

    mbChanged = bNewValue

End Property

Public Property Set GBBitmap(oNewValue As clsGBBitmap)

    Set mGBBitmap = oNewValue

End Property

Public Property Get GBBitmap() As clsGBBitmap

    Set GBBitmap = mGBBitmap

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String
    
    sFilename = msFilename
    
End Property

Private Sub chkNumTiles_Click()

    If chkNumTiles.value = vbChecked Then
        txtNumTiles.Enabled = True
        txtNumTiles.BackColor = &H80000005
        txtNumTiles.SetFocus
    Else
        txtNumTiles.Enabled = False
        txtNumTiles.BackColor = &H8000000F
    End If

End Sub

Private Sub cmdGetWindowsBitmap_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Load a GB Bitmap from a windows .bmp file
'***************************************************************************
    
    With mdiMain.Dialog
        
    'Get filename
        .InitDir = mGBBitmap.intResource_ParentPath & "\Bitmaps"
        .DialogTitle = "Get Windows Bitmap"
        .Filename = ""
        .Filter = "Windows Bitmap (*.bmp)|*.bmp"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
    
    'Load bitmap into the GBBitmap object variable
        Screen.MousePointer = vbHourglass
        
        msBMPfile = .Filename
        
        Dim str As String
        str = mGBBitmap.intResource_ParentPath
        
        Set mGBBitmap = New clsGBBitmap
        mGBBitmap.intResource_ParentPath = str
        
        TileGetterErrorCount = 0
        
        moBMPOffscreen.CreateBitmapFromBMP .Filename
        
        mGBBitmap.GBPalette.GetPalFromBMP moBMPOffscreen
        mGBBitmap.GetBitFromBMP moBMPOffscreen, miPalMap
                
    'Add bitmap to the resource cache
        If mbCached = True Then
            gResourceCache.ReleaseClient Me
        End If
        
        gResourceCache.AddResourceToCache msFilename, mGBBitmap, Me
        mbCached = True
        
    'Set flag for saving
        mbChanged = True
    
    End With

'Update the screen to display the newly loaded bitmap
    
    mGBBitmap.RenderPixels
    
    intResourceClient_Update
    mbChanged = True

    Screen.MousePointer = 0

    If TileGetterErrorCount > 0 Then
        Dim frm As New frmTileGetterErrors
        Set frm.Offscreen = moBMPOffscreen
        Load frm
    End If

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:cmdGetWindowsBitmap_Click Error"
End Sub

Private Sub cmdSave_Click()

'***************************************************************************
'   Save the current bitmap into a .bit file
'***************************************************************************

    On Error GoTo HandleErrors

    PackFile mGBBitmap.intResource_ParentPath & "\Bitmaps\" & GetTruncFilename(msFilename), mGBBitmap
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:cmdSave_Click Error"

End Sub

Private Sub cmdSaveAs_Click()

'***************************************************************************
'   Save the current bitmap under a new filename
'***************************************************************************

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Bitmap"
        .Filename = ""
        .Filter = "GB Bitmap (*.bit)|*.bit"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBBitmap
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBBitmap.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBBitmap, Me
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:cmdSaveAs_Click Error"
End Sub

Private Sub Form_Load()

    On Error GoTo HandleErrors

'***************************************************************************
'   Load the bitmap editor
'***************************************************************************

'Set form's dimensions
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
    With mViewport
        .Zoom = 1
        
        If Not mGBBitmap Is Nothing Then
        
            .ViewportWidth = mGBBitmap.width
            .ViewportHeight = mGBBitmap.height
        
            .SourceWidth = mGBBitmap.width
            .SourceHeight = mGBBitmap.height
        
            .DisplayWidth = mGBBitmap.width
            .DisplayHeight = mGBBitmap.height
            
        End If
        
    End With
    
'Update the visual display
    intResourceClient_Update

'Organize the forms
    CleanUpForms Me

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:Form_Load Error"
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

'***************************************************************************
'   Release memory when closing
'***************************************************************************
    
'Release memory
    
    gResourceCache.ReleaseClient Me
    
    If Not moBMPOffscreen Is Nothing Then
        moBMPOffscreen.Delete
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:Form_Unload Error"
    
End Sub

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors
    
'***************************************************************************
'   Update the visual display
'***************************************************************************
    mGBBitmap.RenderPixels
        
'Display the current bitmap's filename
    Me.Caption = GetTruncFilename(msFilename)

'Clear the display device
    picDisplay.Cls
    
'Draw from the bitmap's offscreen buffer to the display device using a viewport
    If Not mGBBitmap Is Nothing Then
        
        mGBBitmap.Offscreen.Blit picDisplay.hdc, 0, 0
        
        'draw blank tiles
        Dim xWhite As Integer
        Dim yWhite As Integer
        Dim wWhite As Integer
        Dim hWhite As Integer
        
        If mGBBitmap.TileCount > 0 Then
            xWhite = mGBBitmap.TileCount Mod (mGBBitmap.width \ 8)
            yWhite = mGBBitmap.TileCount \ (mGBBitmap.width \ 8)
                
            wWhite = Abs((xWhite * 8) - mGBBitmap.width)
            hWhite = 8
            
            BitBlt picDisplay.hdc, xWhite * 8, yWhite * 8, wWhite, hWhite, picWhite.hdc, 0, 0, vbSrcCopy
        End If
        
    End If
    
'Update all clients of the current resource
    If tType = GB_ACTIVEEDITOR Then
        If Not mGBBitmap Is Nothing Then
            mGBBitmap.UpdateClients Me
            mGBBitmap.UpdateClients Me
        End If
    End If
    
'Refresh the display device to ensure viewing of bitmap
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditBitmap:intResourceClient_Update Error"
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim click As New clsPoint
    
    Select Case gnTool
        Case GB_ZOOM
            
    End Select

End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'picDisplay.MousePointer = vbArrow
    
    Select Case gnTool
        
        Case GB_ZOOM
                        
                    
    End Select
    
End Sub




Private Sub txtNumTiles_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtNumTiles_LostFocus
    End If

End Sub


Private Sub txtNumTiles_LostFocus()
    
    'If val(txtNumTiles.Text) <> mGBBitmap.TileCount Then
    '    Screen.MousePointer = vbHourglass
    '    Set mGBBitmap = New clsGBBitmap
    '    mGBBitmap.TileCount = txtNumTiles.Text
    'Else
    '    Exit Sub
    'End If
    '
    'If msBMPfile = "" Then
    '    MsgBox "You can't use this feature right now.  You must remake the bitmap from the original BMP, then use this feature.", vbInformation, "Uh, No"
    '    chkNumTiles.value = vbUnchecked
    '    txtNumTiles.Text = ""
    '    Screen.MousePointer = 0
    '    Exit Sub
    'End If
    '
    'moBMPOffscreen.CreateBitmapFromBMP msBMPfile
    '
    'mGBBitmap.GBPalette.GetPalFromBMP moBMPOffscreen
    'mGBBitmap.GetBitFromBMP moBMPOffscreen, miPalMap
    '
    'If mbCached = True Then
    '    gResourceCache.ReleaseClient Me
    'End If
    '
    'gResourceCache.AddResourceToCache msFilename, mGBBitmap, Me
    'mbCached = True
    '
    
    mGBBitmap.TileCount = val(txtNumTiles.Text)
    
    mGBBitmap.bClipping = True
    mGBBitmap.RenderPixels
    mGBBitmap.bClipping = False
    
    intResourceClient_Update
    mbChanged = True
    
    'Screen.MousePointer = 0

End Sub


