VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Runner ver 1.0"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   360
      Left            =   3360
      TabIndex        =   11
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Frame fraInput 
      Caption         =   "I&nput"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "*.*"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdBrowseInput 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   585
         Width           =   285
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "Fi&lter:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "&Input Folder:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame fraSourceProgram 
      Caption         =   "Source &Program"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdBrowseSrcProg 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   285
      End
      Begin VB.TextBox txtSourceProgram 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         Caption         =   "Program &Filename:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Index           =   0
      Left            =   6840
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.FileListBox File 
      Height          =   2235
      Index           =   0
      Left            =   6840
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 4755
Private Const DEF_HEIGHT = 4140

Private mnRecCount As Integer
Private msRecPath As String
Private Function mScanPath(sPath As String) As Boolean

    On Error GoTo HandleErrors

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
    
    Dim dirCursor As Integer
    Dim fileCursor As Integer
    
    For dirCursor = Folder(mnRecCount).ListIndex To (Folder(mnRecCount).ListCount - 1)
        
        Dim bScanned As Boolean
        bScanned = mScanPath(Folder(mnRecCount).List(dirCursor))
        
        File(mnRecCount).Path = Folder(mnRecCount).List(dirCursor)
        
        For fileCursor = 0 To (File(mnRecCount).ListCount - 1)
        
            DoEvents
        
            If Not bScanned Then
                
                Dim sFolder As String
                Dim sFile As String
                
                sFolder = Folder(mnRecCount).List(dirCursor) & "\"
                sFile = File(mnRecCount).List(fileCursor)
                
                Shell txtSourceProgram.Text & " " & sFolder & sFile,
                
            End If
            
            mScanPath = True
            
        Next fileCursor
        
    Next dirCursor
    
    On Error Resume Next
    Unload Folder(mnRecCount)
    Unload File(mnRecCount)
    On Error GoTo 0
    mnRecCount = mnRecCount - 1

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMain:mScanPath Error"
End Function

Private Sub cmdBrowseInput_Click()

    Dim sFolder As String
    sFolder = GetPathDialog
    
    If sFolder = "" Then
        Exit Sub
    End If
    
    txtInput.Text = sFolder

End Sub

Private Sub cmdBrowseSrcProg_Click()

    With Dialog
        
        .DefaultExt = "exe"
        .DialogTitle = "Find Source Program"
        .filename = ""
        .Filter = "All Files (*.*)|*.*|Executable Files (*.exe)|*.exe"
        .FilterIndex = 1
        .InitDir = App.Path
        .ShowOpen
        
        If .filename = "" Then
            Exit Sub
        End If
        
        txtSourceProgram.Text = .filename
        
    End With

End Sub

Private Sub cmdExit_Click()

    End

End Sub


Private Sub cmdRun_Click()

    On Error GoTo HandleErrors

    If txtSourceProgram.Text = "" Then
        cmdBrowseSrcProg_Click
        
        If txtSourceProgram.Text = "" Then
            Exit Sub
        End If
        
    End If
    
    If txtInput.Text = "" Then
        cmdBrowseInput_Click
    
        If txtInput.Text = "" Then
            Exit Sub
        End If
    
    End If

    If txtFilter.Text = "" Then
        txtFilter.Text = "*.*"
    End If
    
    File(0).Pattern = txtFilter.Text

    msRecPath = ""
    mnRecCount = 0

    Screen.MousePointer = vbHourglass
    mScanPath txtInput.Text
    Screen.MousePointer = 0
    
    MsgBox "And...your done!", vbInformation, "Deal Wit Me!"
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMain:cmdRun_Click Error"
End Sub

Private Sub Form_Load()

    Me.Width = DEF_WIDTH
    Me.Height = DEF_HEIGHT

End Sub


