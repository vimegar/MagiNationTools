VERSION 5.00
Begin VB.Form frmColFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find/Replace"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmColFindReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Frame fraInput 
      Caption         =   "&Input"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtReplace 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Replace With:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Find Code:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmColFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGBCollisionMap As clsGBCollisionMap

Public Property Get GBCollisionMap() As clsGBCollisionMap

    Set GBCollisionMap = mGBCollisionMap

End Property

Public Property Set GBCollisionMap(oNewValue As clsGBCollisionMap)

    Set mGBCollisionMap = oNewValue

End Property


Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOk_Click()

    If txtFind.Text = "" Then
        MsgBox "You must enter a value to find!", vbInformation, "Input Error"
        txtFind.SetFocus
        Exit Sub
    End If

    If txtReplace.Text = "" Then
        MsgBox "You must enter a value with which to replace!", vbInformation, "Input Error"
        txtReplace.SetFocus
        Exit Sub
    End If

    Dim i As Integer
    
    For i = 0 To (mGBCollisionMap.GBMap.width * mGBCollisionMap.GBMap.height)
        If mGBCollisionMap.CollisionData(i) = val(txtFind.Text) Then
            mGBCollisionMap.CollisionData(i) = val(txtReplace.Text)
        End If
    Next i
    
    Unload Me

End Sub

