VERSION 5.00
Begin VB.Form frmTileGetterErrors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Errors"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmTileGetterErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   619
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   Begin VB.Timer tmrFlash 
      Interval        =   500
      Left            =   6480
      Top             =   8640
   End
   Begin VB.Frame fraErrorList 
      Caption         =   "&Error List"
      Height          =   9015
      Left            =   8280
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   8280
         Width           =   1215
      End
      Begin VB.ListBox lstErrors 
         Height          =   7665
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraKey 
      Caption         =   "&Bitmap w/Errors"
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.PictureBox picDisplay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7680
         Left            =   240
         ScaleHeight     =   7650
         ScaleWidth      =   7650
         TabIndex        =   10
         Top             =   360
         Width           =   7680
      End
      Begin VB.PictureBox picIceCream 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   360
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   3
         Top             =   8655
         Width           =   120
      End
      Begin VB.PictureBox picQuestion 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   360
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   2
         Top             =   8430
         Width           =   120
      End
      Begin VB.PictureBox picYuck 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   360
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   1
         Top             =   8190
         Width           =   120
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "No palette match found"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   8625
         Width           =   1665
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "No pixel match found"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   8385
         Width           =   1500
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Too many colors"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   8145
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmTileGetterErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 11850
Private Const DEF_HEIGHT = 9660

Private moOffscreen As clsOffscreen
Private miErrorMap() As Byte
Private mnErrorX() As Integer
Private mnErrorY() As Integer
Private msName As String
Private msOutputFolder As String

Private mTempOffscreen As New clsOffscreen

Private Sub mDrawErrorBitmap()

    Dim xTile As Integer
    Dim yTile As Integer
    Dim lColor As Long
    
    For yTile = 0 To (moOffscreen.height \ 8) - 1
        For xTile = 0 To (moOffscreen.width \ 8) - 1
            
            BitBlt mTempOffscreen.hdc, xTile * 8, yTile * 8, 8, 8, moOffscreen.hdc, xTile * 8, yTile * 8, vbSrcCopy
            
            If miErrorMap(xTile, yTile) > 0 Then
                
                Select Case miErrorMap(xTile, yTile) - 1
                    Case TooManyColors
                        lColor = vbGreen
                    Case NoPixelMatch
                        lColor = vbRed
                    Case NoPaletteMatch
                        lColor = vbBlue
                End Select
                                
                mTempOffscreen.LineDraw xTile * 8, yTile * 8, (xTile * 8) + 7, (yTile * 8) + 7, lColor, True
                                
            End If
            
        Next xTile
    Next yTile

    StretchBlt picDisplay.hdc, 0, 0, 512, 512, mTempOffscreen.hdc, 0, 0, 256, 256, vbSrcCopy
    picDisplay.Refresh

End Sub

Private Sub mPrintErrors()

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim nFilenum As Integer
    
    nFilenum = FreeFile
    
    Open msOutputFolder & "debug.txt" For Append As #nFilenum
    
    Print #nFilenum, "[" & msName & " Error List]"
    
    For i = 0 To lstErrors.ListCount - 1
        Print #nFilenum, lstErrors.List(i)
    Next i
    
    Print #nFilenum, ""
    Print #nFilenum, ""
    
    Close #nFilenum
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmTileGetterErrors:mPrintErrors Error"
End Sub

Private Sub mUpdate()

    On Error GoTo HandleErrors

    mUpdateErrorList
    mDrawErrorBitmap
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmTileGetterErrors:mUpdate Error"
End Sub

Private Sub mUpdateErrorList()

    Dim i As Integer
    Dim str As String
    
    lstErrors.Clear
    
    ReDim miErrorMap(moOffscreen.width \ 8, moOffscreen.height \ 8)
    ReDim mnErrorX(TileGetterErrorCount)
    ReDim mnErrorY(TileGetterErrorCount)
    
    For i = 1 To TileGetterErrorCount
        With TileGetterErrors(i)
            
            Select Case .nType
                Case TooManyColors
                    str = "There are " & CStr(.ColorCount) & " colors in tile " & Format$(CStr(.X), "00") & ", " & Format$(CStr(.Y), "00") & "."
                Case NoPixelMatch
                    str = "Tile " & Format$(CStr(.X), "00") & ", " & Format$(CStr(.Y), "00") & " does not match the VRAM."
                Case NoPaletteMatch
                    str = "Tile " & Format$(CStr(.X), "00") & ", " & Format$(CStr(.Y), "00") & " does not match a palette."
            End Select
            
            lstErrors.AddItem str
            
            If miErrorMap(.X, .Y) = 0 Then
                miErrorMap(.X, .Y) = .nType + 1
            End If
            
            mnErrorX(i) = .X
            mnErrorY(i) = .Y
            
        End With
    Next i

End Sub

Public Property Set Offscreen(oNewValue As clsOffscreen)

    Set moOffscreen = oNewValue

End Property

Public Property Let sName(sNewValue As String)

    msName = sNewValue

End Property

Public Property Let sOutputFolder(sNewValue As String)

    msOutputFolder = sNewValue

End Property

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT

    If Not gbBatchGoing Then
        MsgBox "The bitmap you just loaded contained one or more errors.  Please correct these errors and try again.", vbCritical, "Tile Errors"
    End If
    
    mTempOffscreen.Create 256, 256
    
    Me.Show
    mUpdate

    mPrintErrors

End Sub
Private Sub Form_Unload(Cancel As Integer)

    mTempOffscreen.Delete
    moOffscreen.Delete

End Sub


Private Sub lstErrors_Click()

    mDrawErrorBitmap

End Sub


Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim xTile As Integer
    Dim yTile As Integer
    
    xTile = X \ 16
    yTile = Y \ 16
    
    Dim i As Integer
    
    For i = 1 To TileGetterErrorCount
        If mnErrorX(i) = xTile And mnErrorY(i) = yTile Then
            lstErrors.ListIndex = i - 1
        End If
    Next i

End Sub


Private Sub tmrFlash_Timer()

    If lstErrors.ListIndex < 0 Then
        Exit Sub
    End If
    
    Dim i As Integer
    
    i = lstErrors.ListIndex + 1
    
    BitBlt picDisplay.hdc, mnErrorX(i) * 16, mnErrorY(i) * 16, 16, 16, picDisplay.hdc, mnErrorX(i) * 16, mnErrorY(i) * 16, vbDstInvert
    picDisplay.Refresh

End Sub


