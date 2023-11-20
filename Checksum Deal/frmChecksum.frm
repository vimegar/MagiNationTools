VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChecksum 
   Caption         =   "Checksum Deal"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2640
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Do Checksum Deal..."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmChecksum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDoIt_Click()

    On Error GoTo HandleErrors

    With Dialog
        .DialogTitle = "Find File"
        .filename = ""
        .Filter = "All Files (*.*)|*.*"
        .InitDir = App.Path
        .ShowOpen
        If .filename = "" Then
            Exit Sub
        End If
        
        Dim sFilename As String
        sFilename = .filename
    End With
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Binary Access Read Write As #nFilenum
        
        Screen.MousePointer = vbHourglass
        Progress.Value = 0
        Progress.Max = LOF(nFilenum)
        
        Do Until EOF(nFilenum)
            
            DoEvents
            
            Dim iByte As Byte
            Get #nFilenum, , iByte
            
            Dim lTotal As Long
            lTotal = lTotal + iByte
            
            Progress.Value = Progress.Value + 1
            
        Loop
        
        
rExit:
    Close #nFilenum

    Screen.MousePointer = 0
    MsgBox "Checksum = " & CStr(lTotal), vbInformation, "Checksum Deal"

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmChecksum:cmdDoIt_Click Error"
    GoTo rExit
End Sub


