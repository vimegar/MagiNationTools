VERSION 5.00
Begin VB.Form frmEditVRAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRAM"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   Icon            =   "frmEditVRAM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   822
   Begin VB.Frame fraEntries 
      Caption         =   "&VRAM Entries"
      Height          =   5655
      Left            =   8280
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkRemap 
         Caption         =   "Re&map Mode"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5160
         Width           =   1815
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
         Height          =   3840
         ItemData        =   "frmEditVRAM.frx":030A
         Left            =   120
         List            =   "frmEditVRAM.frx":030C
         TabIndex        =   1
         Top             =   240
         Width           =   3765
      End
      Begin VB.CommandButton cmdAddVRAMEntry 
         Caption         =   "&Add Bitmap..."
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeleteEntry 
         Caption         =   "&Remove Bitmap"
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmdDown 
         Height          =   1335
         Left            =   3000
         Picture         =   "frmEditVRAM.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdUp 
         Height          =   1335
         Left            =   2040
         Picture         =   "frmEditVRAM.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4200
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Sa&ve As..."
      Height          =   360
      Left            =   10320
      TabIndex        =   7
      Top             =   5760
      Width           =   1830
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   8400
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame fraBank0 
      Caption         =   "Bank 0"
      Height          =   6135
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox picBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5760
         Index           =   1
         Left            =   4200
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   10
         Top             =   240
         Width           =   3840
      End
      Begin VB.PictureBox picBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5760
         Index           =   0
         Left            =   120
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   9
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Bank 1"
         Height          =   195
         Left            =   4200
         TabIndex        =   11
         Top             =   0
         Width           =   510
      End
      Begin VB.Shape shpColors 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   1920
         Index           =   0
         Left            =   3960
         Top             =   240
         Width           =   240
      End
      Begin VB.Shape shpColors 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   1920
         Index           =   1
         Left            =   3960
         Top             =   2160
         Width           =   240
      End
      Begin VB.Shape shpColors 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   1920
         Index           =   2
         Left            =   3960
         Top             =   4080
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEditVRAM"
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

    Private Const DEF_WIDTH = 12420 'In Twips
    Private Const DEF_HEIGHT = 6615 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mViewport As New clsViewport
    Private mSelection As New clsSelection
    
    Private mnSelBank As Byte
    Private msFilename As String
    Private mbChanged As Boolean
    Private mbCached As Boolean
    
'***************************************************************************
'   Resource Object
'***************************************************************************

    Private mGBVRAM As New clsGBVRAM

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Set GBVRAM(oNewValue As clsGBVRAM)

    Set mGBVRAM = oNewValue

End Property

Public Property Get GBVRAM() As clsGBVRAM

    Set GBVRAM = mGBVRAM
    
End Property

Private Sub mPopulateVRAMEntryList(nSelection As Integer)

    On Error GoTo HandleErrors

'***************************************************************************
'   Place info on each VRAM entry into the list box
'***************************************************************************

'Clear the list box before we begin
    lstEntries.Clear
    
    Dim i As Integer
    Dim truncFilename As String
    
'Add the data to the list box with AddItem
    For i = 1 To mGBVRAM.VRAMEntryCount
        truncFilename = GetTruncFilename(mGBVRAM.VRAMEntryFilenames(i))
        If Len(truncFilename) > 24 Then
            truncFilename = Mid$(truncFilename, 1, 24)
        End If
        lstEntries.AddItem Mid$(truncFilename, 1, 24) & String(24 - Len(truncFilename), " ") & " | " & Hex(mGBVRAM.VRAMEntryBaseAddress(i)) & " | " & Format(mGBVRAM.VRAMEntryBank(i), "0")
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
    MsgBox Err.Description, vbCritical, "frmEditVRAM:mPopulateVRAMEntryList Error"
End Sub

Public Property Get SelectedBank() As Byte

'***************************************************************************
'   Retrieve the currently selected bank
'***************************************************************************
    
    SelectedBank = mnSelBank

End Property

Private Sub cmdAddVRAMEntry_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Add a new VRAM entry at address &H8000
'***************************************************************************

'Get filename
    Dim sFilename As String
    
    With mdiMain.Dialog
        .InitDir = mGBVRAM.intResource_ParentPath & "\Bitmaps"
        .DialogTitle = "Load GB Bitmap File"
        .Filename = ""
        .Filter = "GB Bitmap (*.bit)|*.bit"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        sFilename = .Filename
    End With

'Add VRAM entry to VRAM object
    Dim dummy As Integer
    dummy = mGBVRAM.VRAMEntryCount
    
    mGBVRAM.AddVRAMEntry sFilename, 32768, 0
    
    If dummy = mGBVRAM.VRAMEntryCount Then
        sFilename = ""
        Exit Sub
    End If
    
'Update the display
    mPopulateVRAMEntryList lstEntries.ListCount
    
    mbChanged = True
    
    intResourceClient_Update
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdAddVRAMEntry Error"

End Sub

Private Sub cmdDeleteEntry_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Remove the currently selected entry from the VRAM entry list
'***************************************************************************

    Dim curSel As Integer
    
'Check for errors
    If lstEntries.ListCount <= 0 Or lstEntries.ListIndex = -1 Then
        Exit Sub
    End If
    
'Delete the appropriate entry
    curSel = lstEntries.ListIndex + 1

    mGBVRAM.DeleteVRAMEntry curSel
    mPopulateVRAMEntryList curSel
    
'Update the visual
    mbChanged = True
    
    intResourceClient_Update

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdDeleteEntry_Click Error"
End Sub

Private Sub cmdDown_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Re-sort the VRAM entry list by moving the current selection down
'***************************************************************************
    
    Dim BaseAddress(1) As Long
    Dim Filename(1) As String
    Dim bankNum(1) As Byte
    Dim oBitmap(1) As clsGBBitmap
    Dim i As Integer
    Dim curSel As Integer
    
    If lstEntries.ListIndex < 0 Or lstEntries.ListIndex = lstEntries.ListCount - 1 Then
        Exit Sub
    End If
    
    curSel = lstEntries.ListIndex + 1
    
    For i = 0 To 1
        BaseAddress(i) = mGBVRAM.VRAMEntryBaseAddress(i + curSel)
        Filename(i) = mGBVRAM.VRAMEntryFilenames(i + curSel)
        bankNum(i) = mGBVRAM.VRAMEntryBank(i + curSel)
        Set oBitmap(i) = mGBVRAM.VRAMEntryBitmap(i + curSel)
    Next i
    
    mGBVRAM.VRAMEntryBaseAddress(curSel + 1) = BaseAddress(0)
    mGBVRAM.VRAMEntryFilenames(curSel + 1) = Filename(0)
    mGBVRAM.VRAMEntryBank(curSel + 1) = bankNum(0)
    Set mGBVRAM.VRAMEntryBitmap(curSel + 1) = oBitmap(0)
    
    mGBVRAM.VRAMEntryBaseAddress(curSel) = BaseAddress(1)
    mGBVRAM.VRAMEntryFilenames(curSel) = Filename(1)
    mGBVRAM.VRAMEntryBank(curSel) = bankNum(1)
    Set mGBVRAM.VRAMEntryBitmap(curSel) = oBitmap(1)
    
    mPopulateVRAMEntryList curSel
    
    mbChanged = True
    
    intResourceClient_Update
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdDown_Click Error"
End Sub

Private Sub cmdSave_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current VRAM into a .vrm file
'***************************************************************************

    PackFile mGBVRAM.intResource_ParentPath & "\VRAMs\" & GetTruncFilename(msFilename), mGBVRAM
    mbChanged = False

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdSave_Click Error"
End Sub

Private Sub cmdSaveAs_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Save the current VRAM under a new filename
'***************************************************************************

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB VRAM"
        .Filename = ""
        .Filter = "GB VRAMs (*.vrm)|*.vrm"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBVRAM
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBVRAM.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBVRAM, Me
        
        mGBVRAM.OffscreenBank0.Create 128, 192
        mGBVRAM.OffscreenBank1.Create 128, 192
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdSaveAs_Click Error"
End Sub

Private Sub cmdUp_Click()

    On Error GoTo HandleErrors

'***************************************************************************
'   Re-sort the VRAM entry list by moving the current selection up
'***************************************************************************

    Dim BaseAddress(1) As Long
    Dim Filename(1) As String
    Dim bankNum(1) As Byte
    Dim oBitmap(1) As clsGBBitmap
    Dim oBitmapFragments(1) As BITMAP
    Dim i As Integer
    Dim curSel As Integer
    
    If lstEntries.ListIndex <= 0 Then
        Exit Sub
    End If
    
    curSel = lstEntries.ListIndex + 1
    
    For i = 0 To 1
        BaseAddress(i) = mGBVRAM.VRAMEntryBaseAddress(curSel - i)
        Filename(i) = mGBVRAM.VRAMEntryFilenames(curSel - i)
        bankNum(i) = mGBVRAM.VRAMEntryBank(curSel - i)
        Set oBitmap(i) = mGBVRAM.VRAMEntryBitmap(curSel - i)
        'Set oBitmapFragments(i) = mGBVRAM.BitmapFragments(curSel - i)
    Next i
    
    mGBVRAM.VRAMEntryBaseAddress(curSel - 1) = BaseAddress(0)
    mGBVRAM.VRAMEntryFilenames(curSel - 1) = Filename(0)
    mGBVRAM.VRAMEntryBank(curSel - 1) = bankNum(0)
    Set mGBVRAM.VRAMEntryBitmap(curSel - 1) = oBitmap(0)
    'Set mGBVRAM.BitmapFragments(curSel - 1) = oBitmapFragments(0)
    
    mGBVRAM.VRAMEntryBaseAddress(curSel) = BaseAddress(1)
    mGBVRAM.VRAMEntryFilenames(curSel) = Filename(1)
    mGBVRAM.VRAMEntryBank(curSel) = bankNum(1)
    Set mGBVRAM.VRAMEntryBitmap(curSel) = oBitmap(1)
    'Set mGBVRAM.BitmapFragments(curSel) = oBitmapFragments(1)
    
    mPopulateVRAMEntryList curSel - 2
    
    mbChanged = True
    
    intResourceClient_Update

    'For i = 0 To 1
    '    If Not oBitmap(i) Is Nothing Then
    '        oBitmap(i).Delete
    '    End If
    'Next i

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:cmdUp_Click Error"
End Sub

Private Sub Form_Load()

    On Error GoTo HandleErrors

'***************************************************************************
'   Load the VRAM editor
'***************************************************************************

'Set form dimensions
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
'Organize forms
    CleanUpForms Me
    
'Initialize the offscreen buffers
    mGBVRAM.OffscreenBank0.Create 128, 192
    mGBVRAM.OffscreenBank1.Create 128, 192
    
    With mViewport
        .Zoom = 1
        
        .ViewportWidth = 128
        .ViewportHeight = 192
        
        .SourceWidth = 128
        .SourceHeight = 192
        
        .DisplayWidth = 256
        .DisplayHeight = 384
        
    End With

'Draw grid onto display device
    mDrawGridOnBanks
        
'Update the display
    intResourceClient_Update
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:Form_Load Error"
End Sub

Public Property Get bChanged() As Boolean

    bChanged = mbChanged

End Property

Public Property Let bChanged(bNewValue As Boolean)

    mbChanged = bNewValue

End Property

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Select Case gnTool
    
        Case GB_POINTER
        
        Case GB_MARQUEE
        
        Case GB_BRUSH
            
            SelectTool GB_POINTER
            Form_MouseMove Button, Shift, X, Y
        
        Case GB_ZOOM
    
    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Select Case gnTool
    
        Case GB_POINTER
        
            Me.MousePointer = vbArrow
        
        Case GB_MARQUEE
        
            Me.MousePointer = vbArrow
        
        Case GB_BRUSH
            
            Me.MousePointer = vbArrow
        
        Case GB_ZOOM
        
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
    
    On Error GoTo HandleErrors

'***************************************************************************
'   Release memory when form closes
'***************************************************************************
    
    
    gResourceCache.ReleaseClient Me
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:Form_Unload Error"

End Sub

Public Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)
   
    On Error GoTo HandleErrors
 
'***************************************************************************
'   Update the display using the VRAM entry information
'***************************************************************************
    
    Dim Entry As Integer
    Dim gcSrc As New clsPoint
    Dim gcDest As New clsPoint
    Dim copyOK As Boolean

'Setup of bitmap fragments
    mGBVRAM.EnumBitmapFragments

'Draw grid on offscreen banks
    mDrawGridOnBanks
    
'Loop through the VRAM entries and do stuff
    For Entry = 1 To mGBVRAM.VRAMEntryCount
        
        copyOK = True
        
        If mGBVRAM.VRAMEntryBank(Entry) = 0 Then
            
            mGBVRAM.GetPointFromVRAMAddress gcDest, mGBVRAM.VRAMEntryBaseAddress(Entry)
            mGBVRAM.GetPointFromVRAMAddress gcSrc, 0
            
            Dim startX As Integer
            startX = gcDest.X
            Dim nCount As Integer
            
            nCount = 0
            
        'Loop through the VRAM entries and blit them
            While copyOK
                
                If mGBVRAM.VRAMEntryBitmap(Entry).TileCount > 0 Then
                    If nCount = mGBVRAM.VRAMEntryBitmap(Entry).TileCount Then
                        GoTo NextEntry
                    End If
                    nCount = nCount + 1
                End If
    
                If Not mGBVRAM.VRAMEntryBitmap(Entry) Is Nothing Then
                    mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.BlitRect mGBVRAM.OffscreenBank0.hdc, gcDest.X, gcDest.Y, 8, 8, gcSrc.X, gcSrc.Y
                
                'Handle exiting logic
                    Dim tempSrcY As Integer
                    tempSrcY = gcSrc.Y
                    If Not mGBVRAM.GridCursorNext(gcSrc, mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width, mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.height) Then
                        copyOK = False
                    End If
                End If
                
                If gcDest.X + 8 >= 128 Or gcDest.X + 8 >= mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width + startX Then
                    If tempSrcY >= gcSrc.Y Then
                        gcSrc.X = 0
                        gcSrc.Y = gcSrc.Y + 8
                        nCount = nCount + (((startX + mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width) - 128) \ 8)
                    End If
                    gcDest.X = startX
                    gcDest.Y = gcDest.Y + 8
                    If gcDest.Y >= mGBVRAM.OffscreenBank0.height Then
                        copyOK = False
                    End If
                Else
                    If Not mGBVRAM.GridCursorNext(gcDest, mGBVRAM.OffscreenBank0.width, mGBVRAM.OffscreenBank0.height) Then
                        copyOK = False
                    End If
                End If
                
            Wend
            
        Else
                
            mGBVRAM.GetPointFromVRAMAddress gcDest, mGBVRAM.VRAMEntryBaseAddress(Entry)
            mGBVRAM.GetPointFromVRAMAddress gcSrc, 0
            
            startX = gcDest.X
            
            nCount = 0
            
        'Loop through the VRAM entries and blit them
            While copyOK
                
                If mGBVRAM.VRAMEntryBitmap(Entry).TileCount > 0 Then
                    If nCount = mGBVRAM.VRAMEntryBitmap(Entry).TileCount Then
                        GoTo NextEntry
                    End If
                    nCount = nCount + 1
                End If
    
                If Not mGBVRAM.VRAMEntryBitmap(Entry) Is Nothing Then
                    mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.BlitRect mGBVRAM.OffscreenBank1.hdc, gcDest.X, gcDest.Y, 8, 8, gcSrc.X, gcSrc.Y
                
                'Handle exiting logic
                    tempSrcY = gcSrc.Y
                    If Not mGBVRAM.GridCursorNext(gcSrc, mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width, mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.height) Then
                        copyOK = False
                    End If
                End If
                
                If gcDest.X + 8 >= 128 Or gcDest.X + 8 >= mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width + startX Then
                    If tempSrcY >= gcSrc.Y Then
                        gcSrc.X = 0
                        gcSrc.Y = gcSrc.Y + 8
                        nCount = nCount + (((startX + mGBVRAM.VRAMEntryBitmap(Entry).Offscreen.width) - 128) \ 8)
                    End If
                    gcDest.X = startX
                    gcDest.Y = gcDest.Y + 8
                    If gcDest.Y >= mGBVRAM.OffscreenBank1.height Then
                        copyOK = False
                    End If
                Else
                    If Not mGBVRAM.GridCursorNext(gcDest, mGBVRAM.OffscreenBank1.width, mGBVRAM.OffscreenBank1.height) Then
                        copyOK = False
                    End If
                End If
                
            Wend
            
        End If

NextEntry:
    Next Entry

    If gSelection.SrcForm Is Me And (gnTool = GB_BRUSH Or gnTool = GB_MARQUEE) Then
        
    'If using the marquee tool, show the selected area inversed
        If mnSelBank = 0 Then
            mGBVRAM.OffscreenBank0.BlitRaster mGBVRAM.OffscreenBank0.hdc, gSelection.Left * 8, gSelection.Top * 8, gSelection.SelectionWidth * 8, gSelection.SelectionHeight * 8, 0, 0, vbDstInvert
        ElseIf mnSelBank = 1 Then
            mGBVRAM.OffscreenBank1.BlitRaster mGBVRAM.OffscreenBank1.hdc, gSelection.Left * 8, gSelection.Top * 8, gSelection.SelectionWidth * 8, gSelection.SelectionHeight * 8, 0, 0, vbDstInvert
        End If
        
    End If
    
'Update the display
    mViewport.Draw picBank(0).hdc, mGBVRAM.OffscreenBank0
    mViewport.Draw picBank(1).hdc, mGBVRAM.OffscreenBank1
    
    picBank(0).Refresh
    picBank(1).Refresh
    
    mPopulateVRAMEntryList lstEntries.ListIndex
    
'Update all clients of the current VRAM
    If tType = GB_ACTIVEEDITOR Then
        mGBVRAM.UpdateClients Me
        mGBVRAM.UpdateClients Me
    End If
    
'Display filename in the title bar
    Me.Caption = GetTruncFilename(msFilename)

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:intResourceClient_Update Error"
End Sub

Private Sub mDrawGridOnBanks()

    On Error GoTo HandleErrors

'***************************************************************************
'   Draw the grid onto the display buffer
'***************************************************************************

    Dim X As Integer
    Dim Y As Integer
    Dim xLoc As Long
    Dim yLoc As Long
    Dim bgColor As Long
    Dim linColor As Long
        
    bgColor = RGB(113, 64, 99)
    linColor = RGB(196, 141, 180)
        
    For Y = 0 To 23
        For X = 0 To 15
            
            xLoc = X * 8 + 1
            yLoc = Y * 8 + 1
            
            mGBVRAM.OffscreenBank0.RECT xLoc, yLoc, xLoc + 8, yLoc + 8, linColor
            mGBVRAM.OffscreenBank0.RECT xLoc + 1, yLoc + 1, xLoc + 7, yLoc + 7, bgColor
            
            mGBVRAM.OffscreenBank1.RECT xLoc, yLoc, xLoc + 8, yLoc + 8, linColor
            mGBVRAM.OffscreenBank1.RECT xLoc + 1, yLoc + 1, xLoc + 7, yLoc + 7, bgColor
            
        Next X
    Next Y
    
    Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:mDrawGridOnBanks Error"
End Sub

Private Function mGetVRAMAddressFromPoint(click As clsPoint) As Long

    On Error GoTo HandleErrors

'***************************************************************************
'   Returns a VRAM address based on the current click location in pixels
'***************************************************************************

    Set click = mViewport.ScreenClickToGridClick(click.X, click.Y, 8, 8)
    
    mGetVRAMAddressFromPoint = ((click.Y * 256) + click.X * 16) + 32768
        
Exit Function
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:mGetVRAMAddressFromPoint Error"
End Function


Private Sub lstEntries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
        
        Case GB_POINTER
            
        Case GB_MARQUEE
        
        Case GB_BRUSH
        
        Case GB_ZOOM
        
        Case GB_SETTER
        
    End Select

End Sub

Private Sub lstEntries_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Select Case gnTool
        
        Case GB_POINTER
            
            lstEntries.MousePointer = vbArrow
            
        Case GB_MARQUEE
        
            lstEntries.MousePointer = vbCrosshair
            
        Case GB_BRUSH
        
            lstEntries.MousePointer = vbArrow
                        
        Case GB_ZOOM
        
            lstEntries.MousePointer = vbArrow
        
        Case GB_SETTER
        
            lstEntries.MousePointer = vbArrow
        
    End Select

End Sub

Public Sub picBank_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Dim pnt As New clsPoint
    Set pnt = mViewport.ScreenClickToGridClick(X, Y, 8, 8)
    
    Select Case gnTool
        
        Case GB_POINTER
            
            If chkRemap.value = vbChecked Then
                
                chkRemap.value = vbUnchecked
                
                Dim oldpnt As New clsPoint
                Dim oldAddr As Long
                Dim oldBank As Integer
                Dim XOffset As Long
                Dim yOffset As Long
                
                oldAddr = mGBVRAM.VRAMEntryBaseAddress(lstEntries.ListIndex + 1)
                oldBank = mGBVRAM.VRAMEntryBank(lstEntries.ListIndex + 1)
                
                mGBVRAM.GetPointFromVRAMAddress oldpnt, oldAddr
                
                oldpnt.X = oldpnt.X \ 8
                oldpnt.Y = oldpnt.Y \ 8
                
                XOffset = (pnt.X - oldpnt.X)
                yOffset = (pnt.Y - oldpnt.Y)
                
                With gResourceCache

                    Dim i As Integer

                    For i = 1 To .CacheObjectList.count
                        If .CacheObjectList(i).Data.ResourceType = GB_BG Or .CacheObjectList(i).Data.ResourceType = GB_PATTERN Then

                            If MsgBox("Are you sure you want to remap the currently selected bitmap?", vbQuestion Or vbYesNo) = vbNo Then
                                Exit For
                            End If

                            Dim res As New clsGBBackground
                            
                            Set res = .CacheObjectList(i).Data
                            If .CacheObjectList(i).Data.ResourceType = GB_BG Then
                                res.tBackgroundType = GB_RAWBG
                            ElseIf .CacheObjectList(i).Data.ResourceType = GB_PATTERN Then
                                res.tBackgroundType = GB_PATTERNBG
                            End If

                            Dim bx As Integer
                            Dim by As Integer
                            Dim rx As Integer
                            Dim ry As Integer
                            Dim addr As Long
                            
                            Screen.MousePointer = vbHourglass
                            
                            For by = 0 To (mGBVRAM.VRAMEntryBitmap(lstEntries.ListIndex + 1).height \ 8) - 1
                                For bx = 0 To (mGBVRAM.VRAMEntryBitmap(lstEntries.ListIndex + 1).width \ 8) - 1
                                    
                                    For ry = 1 To res.nHeight
                                        For rx = 1 To res.nWidth
                                            
                                            DoEvents
                                    
                                            Dim dumpnt As New clsPoint
                                            dumpnt.X = (oldpnt.X + bx)
                                            dumpnt.Y = (oldpnt.Y + by)
                                            oldAddr = mGBVRAM.GetVRAMAddressFromPoint(dumpnt)
                                            
                                            If (res.VRAMEntryAddress(rx, ry) = oldAddr) And (res.VRAMEntryBank(rx, ry) = oldBank) Then
                                                Dim temppnt As New clsPoint
                                                temppnt.X = dumpnt.X + XOffset
                                                temppnt.Y = dumpnt.Y + yOffset
                                            
                                                res.VRAMEntryAddress(rx, ry) = mGBVRAM.GetVRAMAddressFromPoint(temppnt)
                                                res.VRAMEntryBank(rx, ry) = Index
                                            End If
                                            
                                        Next rx
                                    Next ry
                                    
                                Next bx
                            Next by

                            Screen.MousePointer = 0

                        ElseIf .CacheObjectList(i).Data.ResourceType = GB_SPRITEGROUP Then

                            If MsgBox("Are you sure you want to remap the currently selected bitmap?", vbQuestion Or vbYesNo) = vbNo Then
                                Exit For
                            End If

                            Screen.MousePointer = vbHourglass

                            For by = 0 To (mGBVRAM.VRAMEntryBitmap(lstEntries.ListIndex + 1).height \ 8) - 1
                                For bx = 0 To (mGBVRAM.VRAMEntryBitmap(lstEntries.ListIndex + 1).width \ 8) - 1
                                    
                                    Dim j As Integer
                                    Dim k As Integer
                                    Dim res2 As New clsGBSpriteGroup

                                    Set res2 = .CacheObjectList(i).Data
                                    
                                    For j = 1 To res2.SpriteCount
                                        For k = 1 To res2.Sprites(j).TileCount
                                    
                                            DoEvents
                                    
                                            dumpnt.X = (oldpnt.X + bx)
                                            dumpnt.Y = (oldpnt.Y + by)
                                            oldAddr = mGBVRAM.GetVRAMAddressFromPoint(dumpnt)
                                            
                                            If (res2.Sprites(j).Tiles(k).TileID = (oldAddr - 32768) / 16) And (res2.Sprites(j).Tiles(k).Bank = oldBank) Then
                                                temppnt.X = dumpnt.X + XOffset
                                                temppnt.Y = dumpnt.Y + yOffset
                                            
                                                res2.Sprites(j).Tiles(k).TileID = (mGBVRAM.GetVRAMAddressFromPoint(temppnt) - 32768) / 16
                                                res2.Sprites(j).Tiles(k).Bank = Index
                                            End If
                                            
                                        Next k
                                    Next j
                                    
                                Next bx
                            Next by

                            Screen.MousePointer = 0

'                            Dim j As Integer
'                            Dim k As Integer
'                            Dim res2 As New clsGBSpriteGroup
'
'                            Set res2 = .CacheObjectList(i).Data
'
'                            For j = 1 To res2.SpriteCount
'                                For k = 1 To res2.Sprites(j).TileCount
'                                    addr = (res2.Sprites(j).Tiles(k).TileID * 16) + 32768
'                                    If addr >= oldAddr And addr <= oldAddr + (entryLen * 16) Then
'                                        res2.Sprites(j).Tiles(k).TileID = (addr - 32768 + offset) / 16
'                                        res2.Sprites(j).Tiles(k).Bank = Index
'                                    End If
'                                Next k
'                            Next j
'
                        End If
                    Next i
                End With
            End If
            
        'Move currently selected VRAM entry to the click location in grid units
            If lstEntries.ListIndex < 0 Then
                Exit Sub
            End If
    
            mGBVRAM.VRAMEntryBank(lstEntries.ListIndex + 1) = -(Index <> 0)
    
            pnt.X = X
            pnt.Y = Y
    
            mGBVRAM.VRAMEntryBaseAddress(lstEntries.ListIndex + 1) = mGetVRAMAddressFromPoint(pnt)
            
            mbChanged = True
            
            intResourceClient_Update
            
        Case GB_MARQUEE
        
        'Expand the selection on the form
            mSelection.Left = X \ 16
            mSelection.Top = Y \ 16
            mSelection.Right = mSelection.Left
            mSelection.Bottom = mSelection.Top
            mSelection.AreaWidth = 16
            mSelection.AreaHeight = 24
            mSelection.CellWidth = 8
            mSelection.CellHeight = 8
            
            Set gSelection = mSelection.FixRect
            Set gSelection.SrcForm = Me
            
            mnSelBank = Index
            
            intResourceClient_Update
            
        Case GB_BRUSH
        
            SelectTool GB_POINTER
        
        Case GB_ZOOM
            
            SelectTool GB_POINTER
        
        Case GB_BUCKET
        
            SelectTool GB_POINTER
        
        Case GB_SETTER
        
            SelectTool GB_POINTER
            
        Case GB_REPLACE
            
            SelectTool GB_POINTER
        
    End Select
    
Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:picBank_MouseDown Error"
End Sub

Private Sub picBank_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

    If X >= picBank(Index).width Or X < 0 Or Y >= picBank(Index).height Or Y < 0 Then
        Exit Sub
    End If

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************
    
    Select Case gnTool
        
        Case GB_POINTER
            
            picBank(Index).MousePointer = vbArrow
            
        Case GB_MARQUEE
        
            picBank(Index).MousePointer = vbCrosshair
            
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
                mSelection.Right = X \ 16 + dummyX
                mSelection.Bottom = Y \ 16 + dummyY
                
                Set gSelection = mSelection.FixRect
                Set gSelection.SrcForm = Me
                
                intResourceClient_Update
            
            End If

        Case GB_ZOOM
            
            picBank(Index).MousePointer = vbArrow
        
        Case GB_SETTER
        
            picBank(Index).MousePointer = vbArrow
            
        Case GB_BRUSH
        
            picBank(Index).MousePointer = vbArrow
                        
    End Select

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:picBank_MouseMove Error"

End Sub

Private Sub picBank_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo HandleErrors

'***************************************************************************
'   Standard mouse event logic
'***************************************************************************

    Select Case gnTool
    
        Case GB_POINTER
        
        Case GB_MARQUEE
        
            mnSelBank = Index
    
            SelectTool GB_BRUSH
            Set gSelection.SrcForm = Me
            picBank_MouseMove Index, Button, Shift, X, Y
    
            intResourceClient_Update
        
        Case GB_BRUSH
        
            SelectTool GB_BRUSH
            picBank_MouseMove Index, Button, Shift, X, Y
        
        Case GB_ZOOM
        
    End Select

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditVRAM:picBank_MouseUp Error"

End Sub
