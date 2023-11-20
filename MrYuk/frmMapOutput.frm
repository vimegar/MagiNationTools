VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapOutput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Picture Output"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7755
   Icon            =   "frmMapOutput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   7800
      Width           =   7575
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
         Height          =   825
         Left            =   1320
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Frame fraBatch 
         Caption         =   "&Batch"
         Height          =   855
         Left            =   1425
         TabIndex        =   2
         Top             =   -30
         Width           =   6135
         Begin VB.CommandButton cmdBrowseInput 
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
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtInputDir 
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
            Left            =   480
            TabIndex        =   4
            Top             =   480
            Width           =   4215
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "&Output To..."
            Height          =   360
            Left            =   4815
            TabIndex        =   5
            Top             =   405
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Input Folder:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdLoadMap 
         Caption         =   "&Load Map..."
         Height          =   360
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSaveToBMP 
         Caption         =   "&Save BMP..."
         Enabled         =   0   'False
         Height          =   360
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   1200
      End
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Index           =   0
      Left            =   7800
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.FileListBox File 
      Height          =   2235
      Index           =   0
      Left            =   7800
      Pattern         =   "*.map"
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   360
      Left            =   4920
      TabIndex        =   7
      Top             =   8760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   0
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   6
      Top             =   0
      Width           =   7680
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   8730
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8572
            MinWidth        =   8572
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMapOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements intResourceClient

Private Const DEF_WIDTH = 7845
Private Const DEF_HEIGHT = 9570

Private mGBMap As clsGBMap

Private mnRecCount As Integer
Private msRecPath As String

Private msInputPath As String
Private msOutputPath As String

Private Sub mLoadMap(sFilename As String)

    On Error GoTo HandleErrors

'Release existing resources
    gResourceCache.ReleaseClient Me

'Get resource
    Set mGBMap = gResourceCache.GetResourceFromFile(sFilename, Me)

'Init the pattern resource
    Dim frm As New frmEditBackground
    gbRefOpen = True
    frm.GBBackground.tBackgroundType = GB_PATTERNBG
    frm.bOpening = True
    Set frm.GBBackground = mGBMap.GBBackground
    Load frm
    Unload frm
    gbRefOpen = False
        
    mGBMap.Offscreen.Create mGBMap.width * 16, mGBMap.height * 16
    
'Resize form to fit map aspect ratio
    Dim defSize As Long
    
    defSize = 768
    
    If mGBMap.width = mGBMap.height Then
        mResizeForm defSize, defSize
    Else
        If mGBMap.width > mGBMap.height Then
            mResizeForm defSize, defSize \ (mGBMap.width \ mGBMap.height)
        Else
            mResizeForm defSize \ (mGBMap.height \ mGBMap.width), defSize
        End If
    End If

'Update screen
    intResourceClient_Update

rExit:

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapOutput:mLoadMap Error"
    GoTo rExit
End Sub

Private Sub mResizeForm(BitmapWidth As Long, BitmapHeight As Long)

Exit Sub

    Me.width = (BitmapWidth + 8) * Screen.TwipsPerPixelX
    Me.height = (BitmapHeight + 8 + fraButtons.height + Status.height + 32) * Screen.TwipsPerPixelY

    Dim FormWidth As Long
    Dim FormHeight As Long

    FormWidth = Me.width \ Screen.TwipsPerPixelX
    FormHeight = Me.height \ Screen.TwipsPerPixelY

    picDisplay.width = BitmapWidth
    picDisplay.height = BitmapHeight
    
    fraButtons.Top = picDisplay.height + 8
    
    Progress.Top = Status.Top
    Progress.width = FormWidth - Progress.Left - 8

    Me.Left = (Screen.width - Me.width) \ 2
    Me.Top = (Screen.height - Me.height) \ 2

End Sub


Private Sub cmdAbort_Click()

    gbAbort = True

End Sub

Private Sub cmdBatch_Click()

    On Error GoTo HandleErrors

    If txtInputDir.Text = "" Then
        MsgBox "You must specify an input folder!", vbCritical, "Error"
        Exit Sub
    End If

'Get input path
    msOutputPath = GetPathDialog
    
    If msOutputPath = "" Then
        Exit Sub
    End If
    
    msRecPath = ""
    mnRecCount = 0

    gbAbort = False
    cmdAbort.Visible = True

'Recurse through folders and perform operations
    Screen.MousePointer = vbHourglass
    mOutputFile msInputPath

    If gbAbort Then
        MsgBox "Batch operation aborted!", vbCritical, "Abortion"
    Else
        MsgBox "Batch operation completed successfully!", vbInformation, "Success"
    End If

rExit:
    cmdAbort.Visible = False
    mUpdateStatus ""
    msOutputPath = ""
    Progress.value = 0
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapOutput:cmdBatch_Click Error"
    GoTo rExit
End Sub

Private Function mOutputFile(sPath As String) As Boolean

    Dim dirCursor As Integer
    Dim fileCursor As Integer
    
    If gbAbort Then
        cmdAbort.Visible = False
        Exit Function
    End If
    
    If sPath = msRecPath Then
        Exit Function
    End If
    
    msRecPath = sPath
    
    mnRecCount = mnRecCount + 1
    On Error Resume Next
    Load Folder(mnRecCount)
    Load File(mnRecCount)
    On Error GoTo 0
    Folder(mnRecCount).Path = sPath
    
    If Folder(mnRecCount).ListCount > 0 Then
        Folder(mnRecCount).ListIndex = 0
    End If
    
    For dirCursor = Folder(mnRecCount).ListIndex To (Folder(mnRecCount).ListCount - 1)
        
        Dim bUsed As Boolean
        bUsed = mOutputFile(Folder(mnRecCount).List(dirCursor))
        
        File(mnRecCount).Path = Folder(mnRecCount).List(dirCursor)
        
        For fileCursor = 0 To (File(mnRecCount).ListCount - 1)
        
            DoEvents
        
            If gbAbort Then
                cmdAbort.Visible = False
                Exit Function
            End If
        
            If Not bUsed Then
                Dim i As Integer
                Dim str As String
                Dim sFolder As String
                Dim sFile As String
                
                sFolder = Folder(mnRecCount).List(dirCursor) & "\"
                sFile = File(mnRecCount).List(fileCursor)
                
                mLoadMap sFolder & sFile
                mUpdateStatus sFolder & sFile
                
                For i = Len(sFile) To 1 Step -1
                    If Mid$(sFile, i, 1) = "." Then
                        str = Mid$(sFile, 1, i - 1)
                        Exit For
                    End If
                Next i
                
                SaveToBMP msOutputPath & "\" & str & ".bmp", mGBMap.Offscreen.hdc, BPP24, mGBMap.width * 16, mGBMap.height * 16, Progress
                
            End If
            
            mOutputFile = True
            
        Next fileCursor
        
    Next dirCursor
    
    On Error Resume Next
    Unload Folder(mnRecCount)
    Unload File(mnRecCount)
    On Error GoTo 0
    mnRecCount = mnRecCount - 1
    
End Function

Private Sub cmdBrowseInput_Click()

    msInputPath = UCase$(GetPathDialog)
    
    If msInputPath = "" Then
        msInputPath = txtInputDir.Text
        Exit Sub
    End If
    
    txtInputDir.Text = msInputPath

End Sub

Private Sub cmdLoadMap_Click()

'Determine filename and load map
    With mdiMain.Dialog
        .InitDir = gsProjectPath
        .DialogTitle = "Load Map"
        .Filename = ""
        .Filter = "GB Maps (*.map)|*.map"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        mLoadMap .Filename
        Screen.MousePointer = 0
        
    End With
    
    cmdSaveToBMP.Enabled = True
    
End Sub


Private Sub cmdSaveToBMP_Click()

    On Error GoTo HandleErrors

'Determine filename
    With mdiMain.Dialog
        .InitDir = gsProjectPath
        .DialogTitle = "Save Windows Bitmap"
        .Filename = ""
        .Filter = "Windows Bitmaps (*.bmp)|*.bmp"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        mUpdateStatus .Filename
        SaveToBMP .Filename, mGBMap.Offscreen.hdc, BPP24, mGBMap.width * 16, mGBMap.height * 16, Progress
    
    End With
    
rExit:
    mUpdateStatus ""
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapOutput:cmdSaveToBMP_Click Error"
    GoTo rExit
End Sub

Private Sub mUpdateStatus(sText As String)

    If Len(sText) > 45 Then
        
        Dim i As Integer
        Dim str As String
        
        str = Mid$(sText, Len(sText) - 42)
        
        If InStr(str, "\") <> 0 Then
            For i = 1 To Len(str)
                If Mid$(str, i, 1) = "\" Then
                    str = Mid$(str, i)
                    Exit For
                End If
            Next i
        End If
        
        Status.Panels(1).Text = "..." & str
    Else
        Status.Panels(1).Text = sText
    End If

End Sub

Private Sub Form_Load()

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT

    Progress.value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    gResourceCache.ReleaseClient Me

End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

    Dim X As Integer
    Dim Y As Integer
    Dim patX As Integer
    Dim patY As Integer
    Dim val As Integer
    
    picDisplay.Cls
    mGBMap.Offscreen.Cls
    
    For Y = 0 To (mGBMap.height - 1)
        For X = 0 To (mGBMap.width - 1)
            val = mGBMap.MapData(X + (Y * mGBMap.width) + 1) + 1
            patX = ((val - 1) Mod 16) + 1
            patY = ((val - 1) \ 16) + 1
            If patX >= 0 And patY >= 0 Then
                mGBMap.GBBackground.MapOffscreen.BlitRaster mGBMap.Offscreen.hdc, X * 16, Y * 16, 16, 16, (patX - 1) * 16, (patY - 1) * 16, vbSrcCopy
            End If
        Next X
    Next Y
    
    StretchBlt picDisplay.hdc, 0, 0, picDisplay.width, picDisplay.height, mGBMap.Offscreen.hdc, 0, 0, mGBMap.Offscreen.width, mGBMap.Offscreen.height, vbSrcCopy
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapOutput:intResourceClient_Update Error"
End Sub


