VERSION 5.00
Begin VB.Form frmCollisionCodesProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collision Codes Properties"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmCollisionCodesProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCollisionCodesSrc 
      Caption         =   "&Bitmap Source"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtCollisionCodesSource 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "&Source:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmCollisionCodesProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msFilename As String

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Private Sub cmdBrowse_Click()

    With mdiMain.Dialog
        
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "bmp"
        .DialogTitle = "Load Windows Bitmap"
        .Filename = ""
        .Filter = "Windows Bitmaps (*.bmp)|*.bmp"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
        txtCollisionCodesSource.Text = .Filename
        
    End With

End Sub



Private Sub cmdOK_Click()

    On Error GoTo HandleErrors

    If txtCollisionCodesSource.Text = "(None)" Then
        MsgBox "You must choose a source bitmap!", vbCritical, "Error"
        cmdBrowse.SetFocus
        Exit Sub
    End If

    msFilename = txtCollisionCodesSource.Text

    Me.Hide

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmCollisionCodesProperties:cmdOk_Click Error"

End Sub


Private Sub Form_Load()

    txtCollisionCodesSource.Text = msFilename

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not gbClosingApp

End Sub


