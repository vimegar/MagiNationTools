VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmPalSort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PalSort Test"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picResultPal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.PictureBox picPalBitmap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdRunTest 
      Caption         =   "&Run Test"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmPalSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRunTest_Click()

    With Dialog
        .DialogTitle = "Load Windows BMP"
        .filename = ""
        .Filter = "Windows Bitmaps (*.bmp)|*.bmp"
        .ShowOpen
        If .filename = "" Then
            Exit Sub
        End If
        Dim sFilename As String
        sFilename = .filename
    End With



End Sub

