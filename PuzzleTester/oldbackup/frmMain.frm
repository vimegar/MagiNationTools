VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puzzle Tester"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   707
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTileBitmaps 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   5640
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMaskBitmaps 
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
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picWalkBitmaps 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   4920
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMapBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   4920
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   168
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.PictureBox picBackBuffer 
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
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Timer tmrGameLoop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   3840
   End
   Begin VB.CommandButton cmdCommence 
      Caption         =   "&Load Map.."
      Height          =   360
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
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

Private Const DEF_WIDTH = 6300
Private Const DEF_HEIGHT = 4695







Private Sub cmdCommence_Click()

    tmrGameLoop.Enabled = True

End Sub


Private Sub Form_Load()

    Me.Width = DEF_WIDTH
    Me.Height = DEF_HEIGHT
    
    Me.Show
    glMainHDC = Me.hDC
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub




Private Sub tmrGameLoop_Timer()

    Do While True
        
    'clear windows message buffer
        DoEvents

    'set scrollrate
        Dim ScrollRate As Single
        
        If KeyDown(vbKeyTab) Then
            ScrollRate = PLAYER_RUN_SPEED
            gTony.Speed = RUN
        Else
            ScrollRate = PLAYER_WALK_SPEED
            gTony.Speed = WALK
        End If
    
    'get keystates
        Dim bLeft As Boolean
        Dim bRight As Boolean
        Dim bUp As Boolean
        Dim bDown As Boolean
    
        bLeft = KeyDown(vbKeyLeft)
        bRight = KeyDown(vbKeyRight)
        bUp = KeyDown(vbKeyUp)
        bDown = KeyDown(vbKeyDown)
        
    'set tony dir
        If bUp Then gTony.Dir = DIR_UP
        If bDown Then gTony.Dir = DIR_DOWN
        If bLeft Then gTony.Dir = DIR_LEFT
        If bRight Then gTony.Dir = DIR_RIGHT
        
    'increment tony's frame
        If bUp Or bDown Or bLeft Or bRight Then
            gTony.FrameCursor = gTony.FrameCursor + FRAME_SPEED
            If gTony.FrameCursor >= Len(gTony.FrameList) Then
                gTony.FrameCursor = 0
            End If
        Else
            gTony.FrameCursor = 1
        End If
        
        gTony.Frame = Mid$(gTony.FrameList, Int(gTony.FrameCursor) + 1, 1)
    
    'move tony
        HitDetect gTony, bLeft, bUp, bRight, bDown, ScrollRate
        
    'clear windows message buffer
        DoEvents
    
    'move viewport
        Dim xBound As Integer
        Dim yBound As Integer
        Dim viewportPixelWidth As Integer
        Dim viewportPixelHeight As Integer
        
        viewportPixelWidth = (VIEWPORT_TILE_WIDTH * TILE_PIXEL_WIDTH)
        viewportPixelHeight = (VIEWPORT_TILE_HEIGHT * TILE_PIXEL_HEIGHT)
        
        xBound = viewportPixelWidth / 2
        yBound = viewportPixelHeight / 2
        
        If Abs(gTony.xTile - gViewport.xTile) <= xBound Then
            gViewport.xTile = gViewport.xTile - (ScrollRate * 2)
        End If
        
        If Abs(gTony.xTile - (gViewport.xTile + viewportPixelWidth)) <= xBound Then
            gViewport.xTile = gViewport.xTile + (ScrollRate * 2)
        End If
        
        If Abs(gTony.yTile - gViewport.yTile) <= yBound Then
            gViewport.yTile = gViewport.yTile - (ScrollRate * 2)
        End If
        
        If Abs(gTony.yTile - (gViewport.yTile + viewportPixelHeight)) <= yBound Then
            gViewport.yTile = gViewport.yTile + (ScrollRate * 2)
        End If
        
        If gViewport.xTile < 0 Then gViewport.xTile = 0
        If gViewport.xTile > (gMaps(gnCurrentMap).TileWidth - VIEWPORT_TILE_WIDTH) * TILE_PIXEL_WIDTH Then gViewport.xTile = (gMaps(gnCurrentMap).TileWidth - VIEWPORT_TILE_WIDTH) * TILE_PIXEL_WIDTH
        If gViewport.yTile < 0 Then gViewport.yTile = 0
        If gViewport.yTile > (gMaps(gnCurrentMap).TileHeight - VIEWPORT_TILE_HEIGHT) * TILE_PIXEL_HEIGHT Then gViewport.yTile = (gMaps(gnCurrentMap).TileHeight - VIEWPORT_TILE_HEIGHT) * TILE_PIXEL_HEIGHT
        
    'update display
        'If bLeft Or bRight Or bUp Or bDown Then
            UpdateScreen picDisplay.hDC
            picDisplay.Refresh
        'End If
    Loop

End Sub


