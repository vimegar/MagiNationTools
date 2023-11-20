VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCreature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'general stats
Private msName As String
Private mnLevel As Integer
Private mnExperience As Integer

'battle stats
Private mnEnergy As Integer
Private mnAttack As Integer
Private mnDefense As Integer
Private mnLuck As Integer
Private mnSpecial As Integer
Private mnSpeed As Integer

Public Property Let Level(nNewValue As Integer)

    mnLevel = nNewValue

End Property

Public Property Get Level() As Integer

    Level = mnLevel

End Property

Public Property Get Experience() As Integer

    Experience = mnExperience

End Property

Public Property Let Experience(nNewValue As Integer)

    mnExperience = nNewValue

End Property

Public Property Let Name(sNewValue As String)

    msName = sNewValue

End Property

Public Property Get Name() As String

    Name = msName

End Property

Public Property Let Speed(nNewValue As Integer)

    mnSpeed = nNewValue

End Property

Public Property Get Speed() As Integer

    Speed = mnSpeed

End Property

Public Property Let Special(nNewValue As Integer)

    mnSpecial = nNewValue

End Property

Public Property Get Special() As Integer

    Special = mnSpecial

End Property

Public Property Let Luck(nNewValue As Integer)

    mnLuck = nNewValue

End Property

Public Property Get Luck() As Integer

    Luck = mnLuck

End Property

Public Property Get Defense() As Integer

    Defense = mnDefense

End Property

Public Property Let Defense(nNewValue As Integer)

    mnDefense = nNewValue

End Property

Public Property Get Attack() As Integer

    Attack = mnAttack

End Property

Public Property Let Attack(nNewValue As Integer)

    mnAttack = nNewValue

End Property

Public Property Get Energy() As Integer

    Energy = mnEnergy

End Property

Public Property Let Energy(nNewValue As Integer)

    mnEnergy = nNewValue

End Property
