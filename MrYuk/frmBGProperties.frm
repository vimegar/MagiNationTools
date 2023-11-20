VERSION 5.00
Begin VB.Form frmBGProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BG Properties"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmBGProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1200
   End
End
Attribute VB_Name = "frmBGProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdCancel_Click()

    mbCancel = True
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    On Error GoTo HandleErrors

    If val(txtWidth.Text) < 1 Then
        MsgBox "You must choose a value greater than 0 for width!", vbCritical, "Error"
        txtWidth.SetFocus
        Exit Sub
    End If

    If val(txtHeight.Text) < 1 Then
        MsgBox "You must choose a value greater than 0 for height!", vbCritical, "Error"
        txtHeight.SetFocus
        Exit Sub
    End If

    mnWidth = val(txtWidth.Text)
    mnHeight = val(txtHeight.Text)

    Me.Hide

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmBGProperties:cmdOk_Click Error"
End Sub


Private Sub Form_Load()

    txtWidth.Text = CStr(mnWidth)
    txtHeight.Text = CStr(mnHeight)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not gbClosingApp

End Sub


