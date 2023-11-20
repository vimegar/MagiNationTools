VERSION 5.00
Begin VB.Form frmColorPicker 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Picker"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmColorPicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hsbIntensity 
      Height          =   255
      Left            =   2280
      Max             =   32000
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "16000"
      Top             =   1320
      Value           =   16000
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1920
      TabIndex        =   11
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Timer tmrPalCopy 
      Interval        =   1
      Left            =   120
      Top             =   1320
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Get"
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3240
      TabIndex        =   9
      Top             =   1800
      Width           =   1200
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   255
      Left            =   720
      Max             =   31
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   255
      Left            =   720
      Max             =   31
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   255
      Left            =   720
      Max             =   31
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Intensity:"
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   960
      Width           =   180
   End
   Begin VB.Label lblGreen 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   180
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   180
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Blue:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Green:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Red:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   345
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H80000008&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3240
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mnRed As Integer
Private mnGreen As Integer
Private mnBlue As Integer

Public Property Let nRed(nNewValue As Integer)

    mnRed = nNewValue

End Property

Public Property Let nGreen(nNewValue As Integer)

    mnGreen = nNewValue

End Property

Public Property Let nBlue(nNewValue As Integer)

    mnBlue = nNewValue

End Property

Public Property Get nRed() As Integer

    nRed = mnRed

End Property

Public Property Get nGreen() As Integer

    nGreen = mnGreen

End Property

Public Property Get nBlue() As Integer

    nBlue = mnBlue

End Property

Public Sub Update()

    shpColor.FillColor = RGB(mnRed * 8, mnGreen * 8, mnBlue * 8)
    
    lblRed.Caption = Format$(CStr(mnRed), "00")
    lblGreen.Caption = Format$(CStr(mnGreen), "00")
    lblBlue.Caption = Format$(CStr(mnBlue), "00")

    hsbRed.value = mnRed
    hsbGreen.value = mnGreen
    hsbBlue.value = mnBlue

End Sub

Private Sub chkLock_Click()

    hsbIntensity.value = 31
    hsbIntensity.Tag = 31

End Sub

Private Sub cmdCancel_Click()

    Screen.MousePointer = 0
    Me.Hide
    gnWaitForPal = 0

End Sub

Private Sub cmdCopy_Click()

    Screen.MouseIcon = mdiMain.picDropper.Picture
    Screen.MousePointer = vbCustom
    gnPalCopy = 1

End Sub

Private Sub cmdOk_Click()

    Me.Hide
    gnWaitForPal = 2

End Sub

Private Sub Form_Activate()

    AlwaysOnTop Me, True
    Update

End Sub

Private Sub Form_Load()

    Update

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    gnWaitForPal = 2

End Sub

Private Sub hsbBlue_Change()

    mnBlue = hsbBlue.value
    Update

End Sub

Private Sub hsbBlue_Scroll()

    hsbBlue_Change

End Sub


Private Sub hsbGreen_Change()

    mnGreen = hsbGreen.value
    Update

End Sub

Private Sub hsbGreen_Scroll()

    hsbGreen_Change

End Sub


Private Sub hsbIntensity_Change()

    If hsbIntensity.value > val(hsbIntensity.Tag) Then
        If hsbRed.value < hsbRed.Max Then hsbRed.value = hsbRed.value + 1
        If hsbGreen.value < hsbGreen.Max Then hsbGreen.value = hsbGreen.value + 1
        If hsbBlue.value < hsbBlue.Max Then hsbBlue.value = hsbBlue.value + 1
    ElseIf hsbIntensity.value < val(hsbIntensity.Tag) Then
        If hsbRed.value > 0 Then hsbRed.value = hsbRed.value - 1
        If hsbGreen.value > 0 Then hsbGreen.value = hsbGreen.value - 1
        If hsbBlue.value > 0 Then hsbBlue.value = hsbBlue.value - 1
    End If
    hsbIntensity.Tag = hsbIntensity.value

End Sub

Private Sub hsbIntensity_Scroll()

    hsbIntensity_Change

End Sub


Private Sub hsbRed_Change()

    mnRed = hsbRed.value
    Update

End Sub


Private Sub hsbRed_Scroll()

    hsbRed_Change

End Sub


Private Sub tmrPalCopy_Timer()

    If gnPalCopy = 2 Then
        mnRed = gnPalRed
        mnGreen = gnPalGreen
        mnBlue = gnPalBlue
        Update
        gnPalCopy = 0
    End If

End Sub


