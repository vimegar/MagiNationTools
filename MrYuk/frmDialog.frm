VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDesktop 
      Caption         =   "Desktop"
      Height          =   975
      Left            =   6600
      Picture         =   "frmDialog.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6600
      TabIndex        =   4
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
   Begin VB.DirListBox dirFolder 
      Height          =   3240
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.FileListBox filFile 
      Height          =   3600
      Left            =   3120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Tag             =   "c:"
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msFilenames() As String
Private mnFilenameCount As Integer
Private mbCancel As Boolean

Public Property Get bCancel() As Boolean

    bCancel = mbCancel

End Property


Public Property Let DialogTitle(sValue As String)

    Me.Caption = sValue

End Property


Public Property Get FilenameCount() As Integer

    FilenameCount = mnFilenameCount

End Property


Public Property Get Filenames(Index As Integer) As String

    Filenames = msFilenames(Index)

End Property


Public Property Let InitDir(sValue As String)

    dirFolder.Path = sValue
    
End Property

Private Sub cmdCancel_Click()

    ReDim msFilenames(0)
    mnFilenameCount = 0
    mbCancel = True
    Me.Hide

End Sub

Private Sub cmdDesktop_Click()

    dirFolder.Path = "C:\WINDOWS\Desktop"

End Sub

Private Sub cmdOk_Click()

    On Error GoTo HandleErrors

    Dim i As Integer
    
    mnFilenameCount = 0
    
    For i = 0 To (filFile.ListCount - 1)
        If filFile.Selected(i) Then
            mnFilenameCount = mnFilenameCount + 1
            ReDim Preserve msFilenames(mnFilenameCount)
            msFilenames(mnFilenameCount) = dirFolder.Path & "\" & filFile.List(i)
        End If
    Next i

    If mnFilenameCount = 0 Then
        Exit Sub
    Else
        Me.Hide
    End If

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmDialog:cmdOk_Click Error"
End Sub


Private Sub dirFolder_Change()

    filFile.Path = dirFolder.Path

End Sub

Private Sub drvDrive_Change()

    On Error GoTo HandleErrors
    
    Screen.MousePointer = vbHourglass
    
    dirFolder.Path = drvDrive.Drive
    drvDrive.Tag = drvDrive.Drive
    
    Screen.MousePointer = 0
    
Exit Sub

HandleErrors:
    drvDrive.Drive = drvDrive.Tag
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "File Access Error"
End Sub

Private Sub Form_Load()

    ReDim msFilenames(0)
    mnFilenameCount = 0
    
    mbCancel = False
    
    drvDrive.Drive = "c:"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cmdCancel_Click

End Sub


