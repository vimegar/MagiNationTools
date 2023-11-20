VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'level stats
Private mnLevel As Integer
Private mlExperience As Long

'battle stats
Private mnEnergy As Integer
Private mCreatures As New clsCollection


Public Property Set Creatures(oNewValue As clsCollection)

    Set mCreatures = oNewValue

End Property

Public Property Get Creatures() As clsCollection

    Set Creatures = mCreatures

End Property

Public Property Get Energy() As Integer

    Energy = mnEnergy

End Property

Public Property Let Energy(nNewValue As Integer)

    mnEnergy = nNewValue

End Property

Public Property Get Experience() As Long

    Experience = mlExperience

End Property

Public Property Let Experience(lNewValue As Long)

    mlExperience = lNewValue

End Property

Public Property Get Level() As Integer

    Level = mnLevel

End Property

Public Property Let Level(nNewValue As Integer)

    mnLevel = nNewValue

End Property

