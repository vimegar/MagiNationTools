VERSION 5.00
Begin VB.Form frmGetPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Path"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmGetPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMyDocuments 
      Caption         =   "My Documents"
      Height          =   975
      Left            =   3000
      Picture         =   "frmGetPath.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdDesktop 
      Caption         =   "Desktop"
      Height          =   975
      Left            =   3000
      Picture         =   "frmGetPath.frx":0FCE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1320
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.DirListBox Dir 
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancel As Boolean
Private mPointer As Long
Public Property Get Cancelled() As Boolean

    Cancelled = mbCancel

End Property


Public Property Let Path(sDir As String)

    Dir.Path = sDir

End Property


Public Property Get Path() As String

    Path = Dir.Path

End Property


Private Sub cmdCancel_Click()

    mbCancel = True
    gsCurPath = Dir.Path & "\dummy.d"
    Me.Hide

End Sub

Private Sub cmdDesktop_Click()

    Dir.Path = "C:\Windows\Desktop"

End Sub

Private Sub cmdMyDocuments_Click()

    Dir.Path = "C:\My Documents\"

End Sub


Private Sub cmdOk_Click()

    mbCancel = False
    gsCurPath = Dir.Path & "\dummy.d"
    Me.Hide

End Sub



Private Sub Dir_Change()
    
    Dir.Tag = Dir.Path
    
End Sub

Private Sub Drive_Change()
    
    On Error GoTo HandleErrors
    
    Dir.Path = Drive.Drive
    Drive.Tag = Drive.Drive
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "Access Error"
    Dir.Path = Dir.Tag
    Drive.Drive = Drive.Tag
End Sub


Private Sub Form_Load()

    mPointer = Screen.MousePointer
    Screen.MousePointer = 0
    
    Drive.Tag = "c:\"
    mbCancel = True

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Screen.MousePointer = mPointer

End Sub


