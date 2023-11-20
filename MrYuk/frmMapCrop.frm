VERSION 5.00
Begin VB.Form frmMapCrop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crop Map"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmMapCrop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   840
      TabIndex        =   10
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Frame fraCrop 
      Caption         =   "&Crop Settings"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Height:"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Width:"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Top:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Left:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmMapCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGBMap As clsGBMap
Private mbCancel As Boolean

Public Property Get bCancel() As Boolean

    bCancel = mbCancel
    
End Property


Public Property Set GBMap(oNewValue As clsGBMap)

    Set mGBMap = oNewValue

End Property


Private Sub cmdCancel_Click()

    mbCancel = True
    Unload Me

End Sub


Private Sub cmdOk_Click()

    Dim l As Integer
    Dim t As Integer
    Dim w As Integer
    Dim h As Integer
    
    l = val(txtLeft.Text)
    t = val(txtTop.Text)
    w = val(txtWidth.Text)
    h = val(txtHeight.Text)

    If txtLeft.Text = "" Then
        MsgBox "You must enter a value for Left!", vbInformation, "Data Error"
        txtLeft.SetFocus
        Exit Sub
    End If

    If txtTop.Text = "" Then
        MsgBox "You must enter a value for Top!", vbInformation, "Data Error"
        txtTop.SetFocus
        Exit Sub
    End If

    If w < 10 Then
        MsgBox "You must enter a value greater than or equal to 10 for Width!", vbInformation, "Data Error"
        txtWidth.SetFocus
        Exit Sub
    End If

    If h < 9 Then
        MsgBox "You must enter a value greater than or equal to 9 for Height!", vbInformation, "Data Error"
        txtHeight.SetFocus
        Exit Sub
    End If
        
    Dim X As Integer
    Dim Y As Integer
    Dim srci As Integer
    Dim dsti As Integer
    Dim i As Integer
    
    ReDim iBuffer(w * h) As Byte
    
    dsti = 0
    
    For Y = t To t + h - 1
        For X = l To l + w - 1
        
            srci = X + (Y * mGBMap.width) + 1
            dsti = dsti + 1
            
            If X >= mGBMap.width Or Y >= mGBMap.height Then
                iBuffer(dsti) = 0
            Else
                iBuffer(dsti) = mGBMap.MapData(srci)
            End If
            
        Next X
    Next Y
    
    mGBMap.width = w
    mGBMap.height = h
    
    For Y = 0 To mGBMap.height - 1
        For X = 0 To mGBMap.width - 1
        
            i = X + (Y * mGBMap.width) + 1
            mGBMap.MapData(i) = iBuffer(i)
        
        Next X
    Next Y
    
    mGBMap.Offscreen.Delete
    mGBMap.Offscreen.Create mGBMap.width * 16, mGBMap.height * 16
    
    mbCancel = False
    Unload Me

End Sub


Private Sub Form_Load()

    mbCancel = False

End Sub

