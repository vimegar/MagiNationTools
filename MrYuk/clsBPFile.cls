VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBPFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlSize As Long
Private msFilename As String
Private msFilePath As String

Public Property Let Filename(sNewValue As String)

    msFilename = sNewValue

End Property

Public Property Get Filename() As String

    Filename = msFilename

End Property

Public Property Let FilePath(sNewValue As String)

    msFilePath = sNewValue

End Property

Public Property Get FilePath() As String

    FilePath = msFilePath

End Property

Public Property Let FileSize(lNewValue As Long)

    mlSize = lNewValue

End Property

Public Property Get FileSize() As Long

    FileSize = mlSize

End Property


