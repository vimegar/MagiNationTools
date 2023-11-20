VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load File..."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()

    With Dialog
        .DialogTitle = "Load File"
        .filename = ""
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If .filename = "" Then
            Exit Sub
        End If
        
        Dim t As String
        Dim inputVal As Byte
        Dim nFilenum As Integer
        
        nFilenum = FreeFile
        
        Open .filename For Binary Access Read As #nFilenum
        
        txtText.Text = ""
        
        Do Until EOF(nFilenum)
            Get #nFilenum, , inputVal
            
            t = Hex(inputVal)
            If Len(t) = 1 Then
                t = "0" & t
            End If
            
            txtText.Text = txtText.Text & t & " "
        Loop
        
        Close #nFilenum
        
    End With

End Sub


