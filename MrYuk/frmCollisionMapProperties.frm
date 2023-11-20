VERSION 5.00
Begin VB.Form frmCollisionMapProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collision Map Properties"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmCollisionMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Frame fraMapSrc 
      Caption         =   "&Map Source"
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
      Begin VB.TextBox txtMapSource 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblPattern 
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
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1200
   End
End
Attribute VB_Name = "frmCollisionMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msFilename As String
Private mbCancel As Boolean

Public Property Get bCancel() As Boolean

    bCancel = mbCancel

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Private Sub cmdBrowse_Click()

    With mdiMain.Dialog
        
'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "map"
        .DialogTitle = "Choose map source file"
        .Filename = ""
        .Filter = "GB Maps (*.map)|*.map"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
        txtMapSource.Text = .Filename
        
    End With

End Sub



Private Sub cmdCancel_Click()

    mbCancel = True
    Me.Hide

End Sub

Private Sub cmdOk_Click()
    
    On Error GoTo HandleErrors

    If txtMapSource.Text = "(None)" Then
        MsgBox "You must choose a source map!", vbCritical, "Error"
        cmdBrowse.SetFocus
        Exit Sub
    End If

    msFilename = txtMapSource.Text
    Me.Hide

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmCollisionMapProperties:cmdOk_Click Error"

End Sub


Private Sub Form_Load()

    txtMapSource.Text = msFilename

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not gbClosingApp

End Sub
