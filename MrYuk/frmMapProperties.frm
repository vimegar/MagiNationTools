VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Properties"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Frame fraDimensions 
      Caption         =   "&Size"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         Caption         =   "&Height:"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         Caption         =   "&Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame fraPatternSrc 
      Caption         =   "&Pattern Source"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Height          =   360
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtPatternSource 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblPattern 
         AutoSize        =   -1  'True
         Caption         =   "Sou&rce:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   1200
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msParentPath As String
Private msFilename As String
Private mnWidth As Integer
Private mnHeight As Integer
Private mbCancel As Boolean

Public Property Get bCancel() As Boolean

    bCancel = mbCancel

End Property

Public Property Get nWidth() As Integer

    nWidth = mnWidth

End Property

Public Property Let nWidth(nNewValue As Integer)

    mnWidth = nNewValue

End Property

Public Property Get nHeight() As Integer

    nHeight = mnHeight

End Property

Public Property Let nHeight(nNewValue As Integer)

    mnHeight = nNewValue

End Property

Public Property Let ParentPath(sNewValue As String)

    msParentPath = sNewValue

End Property

Public Property Get ParentPath() As String

    ParentPath = msParentPath

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Private Sub cmdBrowse_Click()

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        
'Get filename
        .InitDir = msParentPath & "\Patterns"
        .DefaultExt = "pat"
        .DialogTitle = "Choose pattern source file"
        .Filename = ""
        .Filter = "GB Patterns (*.pat)|*.pat"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
        txtPatternSource.Text = .Filename
        
    End With

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapProperties:cmdBrowse_Click Error"

End Sub



Private Sub cmdCancel_Click()

    mbCancel = True
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    On Error GoTo HandleErrors

    If txtPatternSource.Text = "(None)" Then
        MsgBox "You must choose a source pattern!", vbCritical, "Error"
        cmdBrowse.SetFocus
        Exit Sub
    End If

    If val(txtWidth.Text) < 10 Then
        MsgBox "You must choose a value greater than or equal to 10 for width!", vbCritical, "Error"
        txtWidth.SetFocus
        Exit Sub
    End If

    If val(txtHeight.Text) < 9 Then
        MsgBox "You must choose a value greater than or equal to 9 for height!", vbCritical, "Error"
        txtHeight.SetFocus
        Exit Sub
    End If

    msFilename = txtPatternSource.Text
    mnWidth = val(txtWidth.Text)
    mnHeight = val(txtHeight.Text)

    mbCancel = False
    Me.Hide

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapProperties:cmdOk_Click Error"
End Sub


Private Sub Form_Load()

    txtWidth.Text = CStr(mnWidth)
    txtHeight.Text = CStr(mnHeight)
    txtPatternSource.Text = msFilename

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not gbClosingApp

End Sub


