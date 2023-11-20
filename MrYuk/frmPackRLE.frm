VERSION 5.00
Begin VB.Form frmPackRLE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RLE Compression"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmPackRLE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMyDocuments 
      Caption         =   "&My Document"
      Height          =   360
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1200
   End
   Begin VB.Frame fraType 
      Caption         =   "&File Type"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   7575
      Begin VB.OptionButton optType 
         Caption         =   "Pa&ttern"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Tag             =   "Pattern"
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Bitmap"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Tag             =   "Bitmap"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optType 
         Caption         =   "Co&llision Map"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Tag             =   "Collision Map"
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Map"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Tag             =   "Map"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Any"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "Any"
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdDesktop 
      Caption         =   "&Desktop"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7800
      TabIndex        =   11
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "Select &None"
      Height          =   360
      Left            =   7800
      TabIndex        =   13
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   360
      Left            =   7800
      TabIndex        =   12
      Top             =   1080
      Width           =   1200
   End
   Begin VB.DirListBox dirFolder 
      Height          =   2340
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "c:"
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdPack 
      Caption         =   "&Pack..."
      Default         =   -1  'True
      Height          =   360
      Left            =   7800
      TabIndex        =   10
      Top             =   120
      Width           =   1200
   End
   Begin VB.FileListBox filList 
      Height          =   79260
      Left            =   5280
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmPackRLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDesktop_Click()

    dirFolder.Path = "c:\windows\desktop"

End Sub

Private Sub cmdMyDocuments_Click()

    dirFolder.Path = "c:\my documents"

End Sub

Private Sub cmdPack_Click()

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim typeIndex As Integer
    Dim nHeaderSize As Integer
    
    For i = 0 To (optType.count - 1)
        If optType(i).value = True Then
            typeIndex = i
            Exit For
        End If
    Next i

    Select Case optType(typeIndex).Tag
        Case "Any"
            nHeaderSize = 0
        Case "Pattern"
            nHeaderSize = 0
        Case "Bitmap"
            nHeaderSize = 0
        Case "Map"
            nHeaderSize = 7
        Case "Collision Map"
            nHeaderSize = 3
    End Select

    Screen.MousePointer = vbHourglass
    
    CreateDir dirFolder.Path & "\RLEOutput"
    
    On Error Resume Next
    Kill dirFolder.Path & "\RLEOutput\*.*"
    On Error GoTo 0
    
    For i = 0 To (filList.ListCount - 1)
        If filList.Selected(i) Then
            PackRLE dirFolder.Path & "\" & filList.List(i), dirFolder.Path & "\RLEOutput\" & filList.List(i), nHeaderSize
        End If
    Next i

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmPackRLE:cmdPack_Click Error"
End Sub


Private Sub cmdSelectAll_Click()

    Dim i As Integer
    
    For i = 0 To (filList.ListCount - 1)
        filList.Selected(i) = True
    Next i

End Sub


Private Sub cmdSelectNone_Click()

    Dim i As Integer
    
    For i = 0 To (filList.ListCount - 1)
        filList.Selected(i) = False
    Next i

End Sub



Private Sub dirFolder_Change()

    filList.Path = dirFolder.Path

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

    drvDrive.Drive = "C:\"

End Sub


