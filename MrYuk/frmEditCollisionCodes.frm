VERSION 5.00
Begin VB.Form frmEditCollisionCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collision Codes"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmEditCollisionCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   3960
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picPic 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   4080
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   3900
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   3960
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   3900
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   0
      ScaleHeight     =   258
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   0
      Top             =   0
      Width           =   3900
   End
End
Attribute VB_Name = "frmEditCollisionCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements intResourceClient

Private Const DEF_WIDTH = 3990
Private Const DEF_HEIGHT = 4365

Private mGBCollisionCodes As New clsGBCollisionCodes
Private mbOpening As Boolean
Private msFilename As String
Private mbChanged As Boolean
Private mbCached As Boolean
Private mnSelectedCode As Integer
Private mnSelX As Integer
Private mnSelY As Integer

Public Property Let bOpening(bNewValue As Boolean)

    mbOpening = bNewValue

End Property


Public Property Set GBCollisionCodes(oNewValue As clsGBCollisionCodes)

    Set mGBCollisionCodes = oNewValue

End Property

Public Property Get GBCollisionCodes() As clsGBCollisionCodes

    Set GBCollisionCodes = mGBCollisionCodes

End Property

Public Property Get MaskHDC() As Long

    MaskHDC = picMask.hdc

End Property

Public Property Get PicHDC() As Long

    PicHDC = picPic.hdc

End Property

Private Sub mCreateMaskFromPicture(hDestDC As Long, ByVal srcPic As Picture)

    On Error GoTo HandleErrors

    Dim X As Integer
    Dim Y As Integer
    Dim hImageDC As Long
    Dim bm As BITMAP
    Dim hBM As Long
    
    hBM = srcPic.Handle
    modWinAPI.GetObject hBM, Len(bm), bm
    hImageDC = CreateCompatibleDC(glMainHDC)
    SelectObject hImageDC, hBM
    
    For Y = 0 To bm.bmHeight - 1
        For X = 0 To bm.bmWidth - 1
            'DoEvents
            If GetPixel(hImageDC, X, Y) <> 0 Then
                SetPixel hDestDC, X, Y, vbBlack
            Else
                SetPixel hDestDC, X, Y, vbWhite
            End If
        Next X
    Next Y

    DeleteDC hImageDC
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:mCreateMaskFromPicture Error"
End Sub

Public Property Get SelectedCode() As Integer

    SelectedCode = mnSelectedCode

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Private Sub cmdSave_Click()

'***************************************************************************
'   Save the current map into a .map file
'***************************************************************************

    On Error GoTo HandleErrors

    PackFile msFilename, mGBCollisionCodes
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:cmdSave_Click Error"

End Sub

Private Sub cmdSaveAs_Click()

'***************************************************************************
'   Save the current collision codes under a new filename
'***************************************************************************

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Collision Codes"
        .Filename = ""
        .Filter = "GB Collision Codes (*.clc)|*.clc"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBCollisionCodes
        msFilename = .Filename
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:cmdSaveAs_Click Error"
End Sub


Private Sub Form_Load()

    On Error GoTo HandleErrors

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
    If Not mbOpening Then
        
        frmCollisionCodesProperties.sFilename = "(None)"
        frmCollisionCodesProperties.Show vbModal
    
        Screen.MousePointer = vbHourglass
        
        mGBCollisionCodes.sBitmapFile = frmCollisionCodesProperties.sFilename
    
        gbClosingApp = True
        Unload frmCollisionCodesProperties
        gbClosingApp = False
    
        Set mGBCollisionCodes.BITMAP = LoadPicture(mGBCollisionCodes.sBitmapFile)
        Screen.MousePointer = 0
        
        mbChanged = True
    
    End If
    
    Screen.MousePointer = vbHourglass
    
    Screen.MousePointer = vbHourglass
    intResourceClient_Update
    mCreateMaskFromPicture picMask.hdc, mGBCollisionCodes.BITMAP
    picPic.Picture = mGBCollisionCodes.BITMAP
    picPic.Refresh
    
    Screen.MousePointer = 0

    intResourceClient_Update
    
    CleanUpForms Me
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:Form_Load Error"
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

    If Not gbClosingApp Then
        Cancel = True
        Me.Hide
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:Form_Unload Error"
End Sub


Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

    'Me.Caption = "collisioncodes.clc" 'GetTruncFilename(msFilename)
    
    picDisplay.Picture = mGBCollisionCodes.BITMAP
    
    picDisplay.Line ((mnSelX - 1) * 16, (mnSelY - 1) * 16)-((mnSelX * 16) - 1, (mnSelY * 16) - 1), vbRed, B
    picDisplay.Refresh

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditCollisionCodes:intResourceClient_Update Error"
End Sub


Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SelectTool GB_POINTER
    
    Set gSelection.SrcForm = Me
    
    mnSelX = (X \ 16) + 1
    mnSelY = (Y \ 16) + 1
    
    mnSelectedCode = (X \ 16) + ((Y \ 16) * 16)
    gnReplaceTile = mnSelectedCode
    
    intResourceClient_Update
    
    SelectTool GB_BRUSH

End Sub


Private Sub tmrHide_Timer()

    tmrHide.Enabled = False
    
    Me.Top = 8
    Me.Left = 8

End Sub


