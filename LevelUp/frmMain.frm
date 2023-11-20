VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Level Up"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   195
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3000
      TabIndex        =   32
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   3000
      TabIndex        =   31
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "20"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Frame fraRes 
      Caption         =   "At Level 00"
      Height          =   2655
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   2655
      Begin VB.Label txtResLuck 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label txtResSpeed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label txtResSpecial 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label txtResDefense 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label txtResAttack 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label txtResEnergy 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblResLuck 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label lblResSpeed 
         AutoSize        =   -1  'True
         Caption         =   "&Speed:"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   960
         Width           =   510
      End
      Begin VB.Label lblResSpecial 
         AutoSize        =   -1  'True
         Caption         =   "Spe&cial:"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblResDefense 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblResAttack 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   510
      End
      Begin VB.Label lblResEnergy 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame fraStartingStats 
      Caption         =   "&Starting Stats"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboRegion 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtLuck 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "3"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "5"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtSpecial 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "4"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "6"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtAttack 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "8"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtEnergy 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "10"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Region:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Luck:"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   13
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Defense:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Spe&cial:"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   570
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Speed:"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Attack:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Energy:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "&Name:"
      Height          =   195
      Index           =   8
      Left            =   3000
      TabIndex        =   33
      Top             =   240
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "&Level:"
      Height          =   195
      Index           =   7
      Left            =   3000
      TabIndex        =   15
      Top             =   1320
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LevelUpNaroom As String
Private LevelUpUnderneath As String
Private LevelUpCald As String
Private LevelUpOrothe As String
Private LevelUpArderial As String
Private EnergyUp(99) As String


Private Sub cmdCalculate_Click()

    Dim i As Integer
    Dim rand As Integer
    Dim resEnergy As Integer
    Dim resAttack As Integer
    Dim resDefense As Integer
    Dim resSpecial As Integer
    Dim resSpeed As Integer
    Dim resLuck As Integer
    Dim level As Integer
    
    level = Val(txtLevel.Text)
    
    resEnergy = Val(txtEnergy.Text)
    resAttack = Val(txtAttack.Text)
    resDefense = Val(txtDefense.Text)
    resSpecial = Val(txtSpecial.Text)
    resSpeed = Val(txtSpeed.Text)
    resLuck = Val(txtLuck.Text)
    
    Select Case cboRegion.Text
        Case "Naroom"
            For i = 2 To level
                resAttack = resAttack + Rnd * (Val(Mid$(LevelUpNaroom, 1, 1)))
                resSpecial = resSpecial + Rnd * (Val(Mid$(LevelUpNaroom, 2, 1)))
                resSpeed = resSpeed + Rnd * (Val(Mid$(LevelUpNaroom, 3, 1)))
                resDefense = resDefense + Rnd * (Val(Mid$(LevelUpNaroom, 4, 1)))
                resLuck = resLuck + Rnd * (Val(Mid$(LevelUpNaroom, 5, 1)))
            Next i
        Case "Underneath"
            For i = 2 To level
                resAttack = resAttack + Rnd * (Val(Mid$(LevelUpUnderneath, 1, 1)))
                resSpecial = resSpecial + Rnd * (Val(Mid$(LevelUpUnderneath, 2, 1)))
                resSpeed = resSpeed + Rnd * (Val(Mid$(LevelUpUnderneath, 3, 1)))
                resDefense = resDefense + Rnd * (Val(Mid$(LevelUpUnderneath, 4, 1)))
                resLuck = resLuck + Rnd * (Val(Mid$(LevelUpUnderneath, 5, 1)))
            Next i
        Case "Cald"
            For i = 2 To level
                resAttack = resAttack + Rnd * (Val(Mid$(LevelUpCald, 1, 1)))
                resSpecial = resSpecial + Rnd * (Val(Mid$(LevelUpCald, 2, 1)))
                resSpeed = resSpeed + Rnd * (Val(Mid$(LevelUpCald, 3, 1)))
                resDefense = resDefense + Rnd * (Val(Mid$(LevelUpCald, 4, 1)))
                resLuck = resLuck + Rnd * (Val(Mid$(LevelUpCald, 5, 1)))
            Next i
        Case "Orothe"
            For i = 2 To level
                resAttack = resAttack + Rnd * (Val(Mid$(LevelUpOrothe, 1, 1)))
                resSpecial = resSpecial + Rnd * (Val(Mid$(LevelUpOrothe, 2, 1)))
                resSpeed = resSpeed + Rnd * (Val(Mid$(LevelUpOrothe, 3, 1)))
                resDefense = resDefense + Rnd * (Val(Mid$(LevelUpOrothe, 4, 1)))
                resLuck = resLuck + Rnd * (Val(Mid$(LevelUpOrothe, 5, 1)))
            Next i
        Case "Arderial"
            For i = 2 To level
                resAttack = resAttack + Rnd * (Val(Mid$(LevelUpArderial, 1, 1)))
                resSpecial = resSpecial + Rnd * (Val(Mid$(LevelUpArderial, 2, 1)))
                resSpeed = resSpeed + Rnd * (Val(Mid$(LevelUpArderial, 3, 1)))
                resDefense = resDefense + Rnd * (Val(Mid$(LevelUpArderial, 4, 1)))
                resLuck = resLuck + Rnd * (Val(Mid$(LevelUpArderial, 5, 1)))
            Next i
    End Select
    
    For i = 2 To level
        rand = Rnd * 3
        resEnergy = resEnergy + (Val(RTrim(Mid$(EnergyUp(i), ((rand * 3 + 1)), 3))))
    Next i
    
    If resEnergy > 99 Then resEnergy = 99
    If resAttack > 99 Then resAttack = 99
    If resSpecial > 99 Then resSpecial = 99
    If resSpeed > 99 Then resSpeed = 99
    If resDefense > 99 Then resDefense = 99
    If resLuck > 99 Then resLuck = 99
    
    fraRes.Caption = "At Level " & Format$(CStr(level), "00")
    
    txtResEnergy.Caption = CStr(resEnergy)
    txtResAttack.Caption = CStr(resAttack)
    txtResSpecial.Caption = CStr(resSpecial)
    txtResSpeed.Caption = CStr(resSpeed)
    txtResDefense.Caption = CStr(resDefense)
    txtResLuck.Caption = CStr(resLuck)

End Sub

Private Sub cmdPrint_Click()

    Me.PrintForm

End Sub


Private Sub Form_Load()

    Randomize Timer

    cboRegion.ListIndex = 0

    'define tables
    
    LevelUpNaroom = "21543"
    LevelUpUnderneath = "32154"
    LevelUpCald = "54321"
    LevelUpOrothe = "15432"
    LevelUpArderial = "43215"
     
    EnergyUp(1!) = "1   1   1   1"
    EnergyUp(2!) = "1   1   1   2"
    EnergyUp(3!) = "1   1   1   2"
    EnergyUp(4!) = "1   1   1   2"
    EnergyUp(5!) = "1   1   1   2"
    EnergyUp(6!) = "1   1   1   2"
    EnergyUp(7!) = "1   1   1   2"
    EnergyUp(8!) = "1   1   1   2"
    EnergyUp(9!) = "1   1   1   2"
    EnergyUp(10) = "1   1   1   2"
    EnergyUp(11) = "1   1   1   2"
    EnergyUp(12) = "1   1   1   2"
    EnergyUp(13) = "1   1   1   2"
    EnergyUp(14) = "1   1   1   2"
    EnergyUp(15) = "1   1   1   2"
    EnergyUp(16) = "1   2   1   2"
    EnergyUp(17) = "1   2   1   2"
    EnergyUp(18) = "1   2   1   2"
    EnergyUp(19) = "1   2   1   2"
    EnergyUp(20) = "1   2   1   2"
    EnergyUp(21) = "1   2   1   2"
    EnergyUp(22) = "1   2   1   2"
    EnergyUp(23) = "1   2   1   2"
    EnergyUp(24) = "1   2   1   2"
    EnergyUp(25) = "1   2   1   2"
    EnergyUp(26) = "1   2   2   2"
    EnergyUp(27) = "1   2   2   2"
    EnergyUp(28) = "1   2   2   2"
    EnergyUp(29) = "1   2   2   2"
    EnergyUp(30) = "1   2   2   2"
    EnergyUp(31) = "1   2   2   2"
    EnergyUp(32) = "1   2   2   3"
    EnergyUp(33) = "1   2   2   3"
    EnergyUp(34) = "1   2   2   3"
    EnergyUp(35) = "1   2   2   3"
    EnergyUp(36) = "1   2   2   3"
    EnergyUp(37) = "1   2   2   3"
    EnergyUp(38) = "2   2   2   3"
    EnergyUp(39) = "2   2   2   3"
    EnergyUp(40) = "2   2   2   3"
    EnergyUp(41) = "2   2   2   3"
    EnergyUp(42) = "2   2   2   3"
    EnergyUp(43) = "2   2   2   3"
    EnergyUp(44) = "2   2   2   3"
    EnergyUp(45) = "2   2   2   3"
    EnergyUp(46) = "2   2   2   3"
    EnergyUp(47) = "2   2   2   3"
    EnergyUp(48) = "2   2   2   3"
    EnergyUp(49) = "2   2   2   3"
    EnergyUp(50) = "2   2   2   3"
    EnergyUp(51) = "2   2   2   3"
    EnergyUp(52) = "2   2   2   3"
    EnergyUp(53) = "2   2   2   3"
    EnergyUp(54) = "2   2   2   3"
    EnergyUp(55) = "2   2   2   3"
    EnergyUp(56) = "2   3   2   3"
    EnergyUp(57) = "2   3   2   3"
    EnergyUp(58) = "2   3   2   3"
    EnergyUp(59) = "2   3   2   3"
    EnergyUp(60) = "2   3   2   3"
    EnergyUp(61) = "2   3   2   3"
    EnergyUp(62) = "2   3   2   3"
    EnergyUp(63) = "2   3   2   3"
    EnergyUp(64) = "2   3   2   4"
    EnergyUp(65) = "2   3   2   4"
    EnergyUp(66) = "2   3   2   4"
    EnergyUp(67) = "2   3   2   4"
    EnergyUp(68) = "2   3   2   4"
    EnergyUp(69) = "2   3   2   4"
    EnergyUp(70) = "2   3   2   4"
    EnergyUp(71) = "2   3   2   4"
    EnergyUp(72) = "2   3   2   4"
    EnergyUp(73) = "2   3   2   4"
    EnergyUp(74) = "2   3   2   4"
    EnergyUp(75) = "2   3   2   4"
    EnergyUp(76) = "2   3   2   4"
    EnergyUp(77) = "2   3   2   4"
    EnergyUp(78) = "2   3   2   4"
    EnergyUp(79) = "2   3   2   4"
    EnergyUp(80) = "2   3   2   4"
    EnergyUp(81) = "2   3   2   4"
    EnergyUp(82) = "2   3   2   4"
    EnergyUp(83) = "2   3   2   4"
    EnergyUp(84) = "2   3   2   4"
    EnergyUp(85) = "2   3   2   4"
    EnergyUp(86) = "2   3   3   4"
    EnergyUp(87) = "2   3   3   4"
    EnergyUp(88) = "2   3   3   4"
    EnergyUp(89) = "2   4   3   4"
    EnergyUp(90) = "2   4   3   4"
    EnergyUp(91) = "2   4   3   4"
    EnergyUp(92) = "2   4   3   4"
    EnergyUp(93) = "2   4   3   4"
    EnergyUp(94) = "2   4   3   4"
    EnergyUp(95) = "2   4   3   4"
    EnergyUp(96) = "2   4   3   4"
    EnergyUp(97) = "2   4   3   4"
    EnergyUp(98) = "2   4   3   4"
    EnergyUp(99) = "3   4   3   4"
        
End Sub


