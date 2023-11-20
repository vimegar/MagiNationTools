VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Stats"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Commence!"
      BeginProperty Font 
         Name            =   "Papyrus LET"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   60
      Top             =   5280
      Width           =   2175
   End
   Begin TabDlg.SSTab TabDialog 
      Height          =   5175
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   9128
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Hero"
      TabPicture(0)   =   "frmStartUp.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraHero"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Creatures"
      TabPicture(1)   =   "frmStartUp.frx":001C
      Tab(1).ControlCount=   56
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMonName(3)"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "txtMonEnergy(3)"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "txtMonDefense(3)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "txtMonAttack(3)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "txtMonSpecial(3)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "txtMonLuck(3)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "txtMonSpeed(3)"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "txtMonName(2)"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "txtMonEnergy(2)"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "txtMonDefense(2)"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "txtMonAttack(2)"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "txtMonSpecial(2)"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "txtMonLuck(2)"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "txtMonSpeed(2)"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "txtMonName(1)"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "txtMonEnergy(1)"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "txtMonDefense(1)"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "txtMonAttack(1)"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "txtMonSpecial(1)"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "txtMonLuck(1)"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "txtMonSpeed(1)"
      Tab(1).Control(20).Enabled=   -1  'True
      Tab(1).Control(21)=   "txtMonName(0)"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "txtMonSpeed(0)"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "txtMonLuck(0)"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "txtMonSpecial(0)"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "txtMonAttack(0)"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "txtMonDefense(0)"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "txtMonEnergy(0)"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "lblLabel(33)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lblLabel(32)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lblLabel(31)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lblLabel(28)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lblLabel(27)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "lblLabel(26)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lblLabel(25)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lblLabel(24)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lblLabel(23)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lblLabel(21)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "lblLabel(20)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "lblLabel(19)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "lblLabel(18)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "lblLabel(17)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "lblLabel(16)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lblLabel(14)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "lblLabel(13)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lblLabel(12)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lblLabel(11)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lblLabel(10)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "lblLabel(9)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lblLabel(30)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "lblLabel(4)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "lblLabel(8)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lblLabel(7)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lblLabel(6)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "lblLabel(5)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lblLabel(2)"
      Tab(1).Control(55).Enabled=   0   'False
      Begin VB.TextBox txtMonName 
         Height          =   285
         Index           =   3
         Left            =   -70320
         TabIndex        =   44
         Top             =   2980
         Width           =   1815
      End
      Begin VB.TextBox txtMonEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -70320
         TabIndex        =   46
         Top             =   3340
         Width           =   495
      End
      Begin VB.TextBox txtMonDefense 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -69000
         TabIndex        =   52
         Top             =   3700
         Width           =   495
      End
      Begin VB.TextBox txtMonAttack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -70320
         TabIndex        =   50
         Top             =   3700
         Width           =   495
      End
      Begin VB.TextBox txtMonSpecial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -69000
         TabIndex        =   56
         Top             =   4060
         Width           =   495
      End
      Begin VB.TextBox txtMonLuck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -70320
         TabIndex        =   54
         Top             =   4060
         Width           =   495
      End
      Begin VB.TextBox txtMonSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -69000
         TabIndex        =   48
         Top             =   3340
         Width           =   495
      End
      Begin VB.TextBox txtMonName 
         Height          =   285
         Index           =   2
         Left            =   -70320
         TabIndex        =   30
         Top             =   820
         Width           =   1815
      End
      Begin VB.TextBox txtMonEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -70320
         TabIndex        =   32
         Top             =   1180
         Width           =   495
      End
      Begin VB.TextBox txtMonDefense 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -69000
         TabIndex        =   38
         Top             =   1540
         Width           =   495
      End
      Begin VB.TextBox txtMonAttack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -70320
         TabIndex        =   36
         Top             =   1540
         Width           =   495
      End
      Begin VB.TextBox txtMonSpecial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -69000
         TabIndex        =   42
         Top             =   1900
         Width           =   495
      End
      Begin VB.TextBox txtMonLuck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -70320
         TabIndex        =   40
         Top             =   1900
         Width           =   495
      End
      Begin VB.TextBox txtMonSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -69000
         TabIndex        =   34
         Top             =   1180
         Width           =   495
      End
      Begin VB.TextBox txtMonName 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   16
         Top             =   2980
         Width           =   1815
      End
      Begin VB.TextBox txtMonEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   18
         Top             =   3340
         Width           =   495
      End
      Begin VB.TextBox txtMonDefense 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -72600
         TabIndex        =   24
         Top             =   3700
         Width           =   495
      End
      Begin VB.TextBox txtMonAttack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   22
         Top             =   3700
         Width           =   495
      End
      Begin VB.TextBox txtMonSpecial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -72600
         TabIndex        =   28
         Top             =   4060
         Width           =   495
      End
      Begin VB.TextBox txtMonLuck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   26
         Top             =   4060
         Width           =   495
      End
      Begin VB.TextBox txtMonSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -72600
         TabIndex        =   20
         Top             =   3340
         Width           =   495
      End
      Begin VB.TextBox txtMonName 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   2
         Top             =   820
         Width           =   1815
      End
      Begin VB.TextBox txtMonSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -72600
         TabIndex        =   6
         Top             =   1180
         Width           =   495
      End
      Begin VB.TextBox txtMonLuck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   12
         Top             =   1900
         Width           =   495
      End
      Begin VB.TextBox txtMonSpecial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -72600
         TabIndex        =   14
         Top             =   1900
         Width           =   495
      End
      Begin VB.TextBox txtMonAttack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   8
         Top             =   1540
         Width           =   495
      End
      Begin VB.TextBox txtMonDefense 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -72600
         TabIndex        =   10
         Top             =   1540
         Width           =   495
      End
      Begin VB.TextBox txtMonEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   4
         Top             =   1180
         Width           =   495
      End
      Begin VB.Frame fraHero 
         Caption         =   "&Hero"
         Height          =   4575
         Left            =   120
         TabIndex        =   0
         Top             =   460
         Width           =   6855
         Begin VB.TextBox txtHeroEnergy 
            Height          =   285
            Left            =   2040
            TabIndex        =   58
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Starting E&nergy (Level 1):"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Width           =   1785
         End
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   33
         Left            =   -70920
         TabIndex        =   43
         Top             =   2980
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Index           =   32
         Left            =   -70920
         TabIndex        =   45
         Top             =   3340
         Width           =   540
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Index           =   31
         Left            =   -69720
         TabIndex        =   51
         Top             =   3700
         Width           =   645
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Index           =   28
         Left            =   -70920
         TabIndex        =   49
         Top             =   3700
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Special:"
         Height          =   195
         Index           =   27
         Left            =   -69720
         TabIndex        =   55
         Top             =   4060
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Index           =   26
         Left            =   -70920
         TabIndex        =   53
         Top             =   4060
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "S&peed:"
         Height          =   195
         Index           =   25
         Left            =   -69720
         TabIndex        =   47
         Top             =   3340
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   24
         Left            =   -70920
         TabIndex        =   29
         Top             =   820
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Index           =   23
         Left            =   -70920
         TabIndex        =   31
         Top             =   1180
         Width           =   540
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Index           =   21
         Left            =   -69720
         TabIndex        =   37
         Top             =   1540
         Width           =   645
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Index           =   20
         Left            =   -70920
         TabIndex        =   35
         Top             =   1540
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Special:"
         Height          =   195
         Index           =   19
         Left            =   -69720
         TabIndex        =   41
         Top             =   1900
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Index           =   18
         Left            =   -70920
         TabIndex        =   39
         Top             =   1900
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "S&peed:"
         Height          =   195
         Index           =   17
         Left            =   -69720
         TabIndex        =   33
         Top             =   1180
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   16
         Left            =   -74520
         TabIndex        =   15
         Top             =   2980
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Index           =   14
         Left            =   -74520
         TabIndex        =   17
         Top             =   3340
         Width           =   540
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Index           =   13
         Left            =   -73320
         TabIndex        =   23
         Top             =   3700
         Width           =   645
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Index           =   12
         Left            =   -74520
         TabIndex        =   21
         Top             =   3700
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Special:"
         Height          =   195
         Index           =   11
         Left            =   -73320
         TabIndex        =   27
         Top             =   4060
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Index           =   10
         Left            =   -74520
         TabIndex        =   25
         Top             =   4060
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "S&peed:"
         Height          =   195
         Index           =   9
         Left            =   -73320
         TabIndex        =   19
         Top             =   3340
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Index           =   30
         Left            =   -74520
         TabIndex        =   1
         Top             =   820
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "S&peed:"
         Height          =   195
         Index           =   4
         Left            =   -73320
         TabIndex        =   5
         Top             =   1180
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   11
         Top             =   1900
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Special:"
         Height          =   195
         Index           =   7
         Left            =   -73320
         TabIndex        =   13
         Top             =   1900
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Index           =   6
         Left            =   -74520
         TabIndex        =   7
         Top             =   1540
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Index           =   5
         Left            =   -73320
         TabIndex        =   9
         Top             =   1540
         Width           =   645
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   3
         Top             =   1180
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 7215
Private Const DEF_HEIGHT = 6315

Private mbNeedUpdate As Boolean
Private Sub mUpdateCreatures()

    If mbNeedUpdate = False Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 3
        
        mbNeedUpdate = False
        
        'update variables
        With gCreatures(i)
            .Name = txtMonName(i).Text
            .Attack = Val(txtMonAttack(i).Text)
            .Defense = Val(txtMonDefense(i).Text)
            .Energy = Val(txtMonEnergy(i).Text)
            .Luck = Val(txtMonLuck(i).Text)
            .Special = Val(txtMonSpecial(i).Text)
            .Speed = Val(txtMonSpeed(i).Text)
        End With
        
        mbNeedUpdate = True
        
    Next i

End Sub





Private Sub cmdStart_Click()

    frmBattle.Show
    Me.Hide

End Sub


Private Sub Form_Load()

    Me.Width = DEF_WIDTH
    Me.Height = DEF_HEIGHT

    Dim i As Integer

    For i = 0 To 3
        
        mbNeedUpdate = False
        
        'get creature data
        With gCreatures(i)
            .Name = GetIniData("Test Battle", "MonName" & Format$(CStr(i), "0"), App.Path & "\testbat.ini")
            .Attack = Val(GetIniData("Test Battle", "MonAttack" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
            .Defense = Val(GetIniData("Test Battle", "MonDefense" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
            .Energy = Val(GetIniData("Test Battle", "MonEnergy" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
            .Luck = Val(GetIniData("Test Battle", "MonLuck" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
            .Special = Val(GetIniData("Test Battle", "MonSpecial" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
            .Speed = Val(GetIniData("Test Battle", "MonSpeed" & Format$(CStr(i), "0"), App.Path & "\testbat.ini"))
        
            'update text boxes
            txtMonName(i).Text = .Name
            txtMonAttack(i).Text = .Attack
            txtMonDefense(i).Text = .Defense
            txtMonEnergy(i).Text = .Energy
            txtMonLuck(i).Text = .Luck
            txtMonSpecial(i).Text = .Special
            txtMonSpeed(i).Text = .Speed
            
        End With
        
        mbNeedUpdate = True
        
    Next i
    
    mUpdateCreatures
    
    'get hero data
    gHero.Level = Val(GetIniData("Test Battle", "HeroLevel", App.Path & "\testbat.ini"))
    gHero.Energy = Val(GetIniData("Test Battle", "HeroEnergy", App.Path & "\testbat.ini"))
        
    'update text boxes
    txtHeroEnergy.Text = gHero.Energy
        
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Dim i As Integer

    For i = 0 To 3
        
        'save creature data
        With gCreatures(i)
            WriteIniData "Test Battle", "MonName" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Name)
            WriteIniData "Test Battle", "MonAttack" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Attack)
            WriteIniData "Test Battle", "MonDefense" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Defense)
            WriteIniData "Test Battle", "MonEnergy" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Energy)
            WriteIniData "Test Battle", "MonLuck" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Luck)
            WriteIniData "Test Battle", "MonSpecial" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Special)
            WriteIniData "Test Battle", "MonSpeed" & Format$(CStr(i), "0"), App.Path & "\testbat.ini", CStr(.Speed)
        End With
        
    Next i
    
    'save hero data
    WriteIniData "Test Battle", "HeroLevel", App.Path & "\testbat.ini", CStr(gHero.Level)
    WriteIniData "Test Battle", "HeroEnergy", App.Path & "\testbat.ini", CStr(gHero.Energy)

End Sub


Private Sub TabDialog_DblClick()

    mUpdateCreatures

End Sub



Private Sub txtHeroEnergy_Change()

    gHero.Energy = Val(txtHeroEnergy.Text)

End Sub

Private Sub txtMonAttack_Change(Index As Integer)

    mUpdateCreatures

End Sub

Private Sub txtMonDefense_Change(Index As Integer)

    mUpdateCreatures

End Sub

Private Sub txtMonEnergy_Change(Index As Integer)

    mUpdateCreatures

End Sub


Private Sub txtMonLuck_Change(Index As Integer)

    mUpdateCreatures

End Sub

Private Sub txtMonName_Change(Index As Integer)

    mUpdateCreatures
    
End Sub


Private Sub txtMonSpecial_Change(Index As Integer)

    mUpdateCreatures

End Sub

Private Sub txtMonSpeed_Change(Index As Integer)

    mUpdateCreatures

End Sub


