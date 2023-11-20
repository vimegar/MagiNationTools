VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moItems() As Object

Public Sub Add(Item As Object)

    ReDim Preserve moItems(UBound(moItems) + 1)
    
    Set moItems(UBound(moItems)) = Item

End Sub

Public Property Get Count() As Integer

    Count = UBound(moItems) + 1

End Property


Public Property Get Item(Index As Integer) As Object

    Set Item = moItems(Index)

End Property


Public Sub Remove(Item As Object)

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To UBound(moItems)
        If moItems(i) Is Item Then
            For j = i To UBound(moItems) - 1
                Set moItems(j) = moItems(j + 1)
            Next j
            ReDim Preserve moItems(UBound(moItems) - 1)
            Exit For
        End If
    Next i

End Sub

Private Sub Class_Initialize()

    ReDim moItems(0)

End Sub


