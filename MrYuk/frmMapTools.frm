VERSION 5.00
Begin VB.Form frmMapTools 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2700
   Icon            =   "frmMapTools.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Tools"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   3
      Left            =   360
      Picture         =   "frmMapTools.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   9
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   8
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   7
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   5
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   4
      Left            =   0
      Picture         =   "frmMapTools.frx":097E
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   2
      Left            =   0
      Picture         =   "frmMapTools.frx":1022
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   1
      Left            =   360
      Picture         =   "frmMapTools.frx":1696
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Selection Tool"
      Top             =   0
      Width           =   360
   End
   Begin VB.OptionButton optTools 
      Height          =   360
      Index           =   0
      Left            =   0
      Picture         =   "frmMapTools.frx":1D9A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Pointer Tool"
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmMapTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 810 '810
Private Const DEF_HEIGHT = 1050 '2165

Private Sub Form_Load()

    On Error GoTo HandleErrors
    
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
    Me.Show
    
    CleanUpForms Nothing
    
    AlwaysOnTop Me, True

Exit Sub
                                 
HandleErrors:
    MsgBox Err.Description, vbCritical, "frmMapTools:Form_Load Error"
End Sub









Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not gbClosingApp

End Sub


Private Sub optTools_Click(Index As Integer)


End Sub







