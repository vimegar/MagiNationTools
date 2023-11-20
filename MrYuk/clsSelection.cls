VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mnTop As Integer
Private mnLeft As Integer
Private mnBottom As Integer
Private mnRight As Integer

Private mnCellWidth As Integer
Private mnCellHeight As Integer
Private mnAreaWidth As Integer
Private mnAreaHeight As Integer

Private mnCursorX As Integer
Private mnCursorY As Integer

Private mSrcForm As Form
Private mDestForm As Form



Public Property Get DestForm() As Form

    Set DestForm = mDestForm

End Property

Public Property Set DestForm(oNewValue As Form)

    Set mDestForm = oNewValue

End Property

Public Property Get SrcForm() As Form

    Set SrcForm = mSrcForm

End Property

Public Property Set SrcForm(oNewValue As Form)

    Set mSrcForm = oNewValue

End Property

Public Property Let AreaWidth(nNewValue As Integer)

    mnAreaWidth = nNewValue

End Property

Public Property Let AreaHeight(nNewValue As Integer)

    mnAreaHeight = nNewValue

End Property

Public Property Get AreaWidth() As Integer

    AreaWidth = mnAreaWidth

End Property

Public Property Get AreaHeight() As Integer

    AreaHeight = mnAreaHeight

End Property

Public Property Let CellWidth(nNewValue As Integer)

    mnCellWidth = nNewValue

End Property

Public Property Let CellHeight(nNewValue As Integer)

    mnCellHeight = nNewValue

End Property

Public Property Get CellWidth() As Integer

    CellWidth = mnCellWidth

End Property

Public Property Get CellHeight() As Integer

    CellHeight = mnCellHeight

End Property

Public Property Get CursorX() As Integer

    CursorX = mnCursorX

End Property

Public Property Get CursorY() As Integer

    CursorY = mnCursorY

End Property

Public Function GetNextElement() As Integer
    
'***************************************************************************
'   Set the cursor position on the current selection space to the next element
'***************************************************************************
    
    mnCursorX = mnCursorX + 1
    
    If mnCursorX >= SelectionWidth Then
        mnCursorX = 0
        mnCursorY = mnCursorY + 1
    End If
    
    If mnCursorY >= SelectionHeight Then
        GetNextElement = -1
        Exit Function
    End If
    
    GetNextElement = (Left + mnCursorX) + ((Top + mnCursorY) * mnAreaWidth) + 1
    
End Function

Public Property Get SelectionWidth() As Integer
    
    SelectionWidth = Abs(Left - Right)

End Property

Public Property Get SelectionHeight() As Integer
    
    SelectionHeight = Abs(Top - Bottom)

End Property

Public Property Get Top() As Integer

    Top = mnTop

End Property

Public Property Let Top(nNewValue As Integer)

    mnTop = nNewValue

End Property

Public Property Get Left() As Integer

    Left = mnLeft

End Property

Public Property Let Left(nNewValue As Integer)

    mnLeft = nNewValue

End Property

Public Property Get Right() As Integer

    Right = mnRight

End Property

Public Property Let Right(nNewValue As Integer)

    mnRight = nNewValue

End Property

Public Property Get Bottom() As Integer

    Bottom = mnBottom

End Property

Public Property Let Bottom(nNewValue As Integer)

    mnBottom = nNewValue

End Property

Public Function FixRect() As clsSelection

'***************************************************************************
'   Fixes the selection space when the number have become negative
'***************************************************************************

    Dim dummy As Integer
    
    Set FixRect = New clsSelection
    
    FixRect.Top = mnTop
    FixRect.Left = mnLeft
    FixRect.Right = mnRight
    FixRect.Bottom = mnBottom
    FixRect.AreaWidth = mnAreaWidth
    FixRect.AreaHeight = mnAreaHeight
    FixRect.CellWidth = mnCellWidth
    FixRect.CellHeight = mnCellHeight
    
    If FixRect.Right < FixRect.Left Then
        dummy = FixRect.Left
        FixRect.Left = FixRect.Right
        FixRect.Right = dummy
    End If
    
    If FixRect.Bottom < FixRect.Top Then
        dummy = FixRect.Top
        FixRect.Top = FixRect.Bottom
        FixRect.Bottom = dummy
    End If

End Function

Public Function GetFirstElement() As Integer

'***************************************************************************
'   Set the cursor for the selection space to the first element
'***************************************************************************

    mnCursorX = 0
    mnCursorY = 0
    
    GetFirstElement = Left + (Top * mnAreaWidth) + 1

End Function

Private Sub Class_Initialize()

'***************************************************************************
'   Setup intial values
'***************************************************************************

    mnCursorX = -1
    mnCursorY = -1
    mnCellWidth = 1
    mnCellHeight = 1

End Sub
