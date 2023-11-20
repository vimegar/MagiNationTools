VERSION 5.00
Begin VB.Form frmBattle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Battle"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmBattle.frx":0000
   ScaleHeight     =   629
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCreatures 
      Caption         =   "Creatures"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   5175
      Begin VB.CommandButton cmdMonSpecialAttack 
         Caption         =   "Special Attack"
         BeginProperty Font 
            Name            =   "Papyrus LET"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdMonBasicAttack 
         Caption         =   "Basic Attack"
         BeginProperty Font 
            Name            =   "Papyrus LET"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraHeroCommands 
      Caption         =   "Hero"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   5175
   End
   Begin VB.Frame fraStats 
      Caption         =   "Stats"
      Height          =   7215
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdEnemyCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   3480
      Picture         =   "frmBattle.frx":087A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   720
   End
   Begin VB.CommandButton cmdEnemyCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   2
      Left            =   2520
      Picture         =   "frmBattle.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   720
   End
   Begin VB.CommandButton cmdEnemyCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   1
      Left            =   1560
      Picture         =   "frmBattle.frx":0A6C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   720
   End
   Begin VB.CommandButton cmdEnemyCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   720
      Picture         =   "frmBattle.frx":0B22
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   720
   End
   Begin VB.CommandButton cmdPlayerCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   3600
      Picture         =   "frmBattle.frx":0BFC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   720
   End
   Begin VB.CommandButton cmdPlayerCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   2
      Left            =   2640
      Picture         =   "frmBattle.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   720
   End
   Begin VB.CommandButton cmdPlayerCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   1
      Left            =   1680
      Picture         =   "frmBattle.frx":1382
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   720
   End
   Begin VB.CommandButton cmdPlayerCreature 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   720
      Picture         =   "frmBattle.frx":1A2E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   720
   End
   Begin VB.CommandButton cmdHero 
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2160
      Picture         =   "frmBattle.frx":1B46
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   720
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

