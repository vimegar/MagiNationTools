VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puzzle Tester"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPlayerMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   5640
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   4920
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   5280
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim bLeft As Boolean
    Dim bUp As Boolean
    Dim bRight As Boolean
    Dim bDown As Boolean
    
'get keystates
    bLeft = KeyDown(vbKeyLeft)
    bUp = KeyDown(vbKeyUp)
    bRight = KeyDown(vbKeyRight)
    bDown = KeyDown(vbKeyDown)

'set tony's direction
    If bLeft Then gPlayer.Direction = DIR_LEFT
    If bUp Then gPlayer.Direction = DIR_UP
    If bRight Then gPlayer.Direction = DIR_RIGHT
    If bDown Then gPlayer.Direction = DIR_DOWN

'check for hotspots before move
    Dim i As Integer
    
    If bLeft Or bUp Or bRight Or bDown Then
        For i = 0 To (CurrentMap.HotspotCount - 1)
            If (CurrentMap.Hotspots(i).xInTiles = gPlayer.xInTiles) And (CurrentMap.Hotspots(i).yInTiles = gPlayer.yInTiles) And (CurrentMap.Hotspots(i).Direction = gPlayer.Direction) Then
                RunHotspot CurrentMap.MapName, CurrentMap.Hotspots(i)
                Exit Sub
            End If
        Next i
    End If

'move tony
    Dim xNew As Integer
    Dim yNew As Integer

    xNew = gPlayer.xInTiles + bLeft - bRight
    yNew = gPlayer.yInTiles + bUp - bDown

    If gColCodes(CurrentMap.MapData(xNew, yNew)).Walkable = True Then
        gPlayer.xInTiles = gPlayer.xInTiles + bLeft - bRight
        gPlayer.yInTiles = gPlayer.yInTiles + bUp - bDown
    End If

'check for switch flipping
    Dim xCheck As Integer
    Dim yCheck As Integer
    
    xCheck = gPlayer.xInTiles + (gPlayer.Direction = DIR_LEFT) - (gPlayer.Direction = DIR_RIGHT)
    yCheck = gPlayer.yInTiles + (gPlayer.Direction = DIR_UP) - (gPlayer.Direction = DIR_DOWN)
    
    If KeyDown(vbKeySpace) Then
        For i = 0 To (CurrentMap.SwitchCount - 1)
            If (CurrentMap.Switches(i).xInTiles = xCheck) And (CurrentMap.Switches(i).yInTiles = yCheck) Then
                RunSwitch CurrentMap.MapName, CurrentMap.Switches(i)
                Exit For
            End If
        Next i
    End If

'update display
    UpdateScreen
    
'check for hotspots on new tile
    If bLeft Or bUp Or bRight Or bDown Then
        For i = 0 To (CurrentMap.HotspotCount - 1)
            If (CurrentMap.Hotspots(i).xInTiles = gPlayer.xInTiles) And (CurrentMap.Hotspots(i).yInTiles = gPlayer.yInTiles) And (CurrentMap.Hotspots(i).Direction = gPlayer.Direction) Then
                RunHotspot CurrentMap.MapName, CurrentMap.Hotspots(i)
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub


