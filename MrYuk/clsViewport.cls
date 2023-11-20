VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlZoomFactor As Long

Private mnSourceWidth As Integer
Private mnSourceHeight As Integer

Private mnDisplayWidth As Integer
Private mnDisplayHeight As Integer

Private mnViewportX As Integer
Private mnViewportY As Integer
Private mnViewportWidth As Integer
Private mnViewportHeight As Integer

Public Function BoundCheck(ByVal click As clsPoint) As clsPoint

    If click.X < 0 Then
        click.X = 0
    End If
            
    If click.X > mnSourceWidth - (mnViewportWidth \ mlZoomFactor) Then
        click.X = mnSourceWidth - (mnViewportWidth \ mlZoomFactor)
    End If
            
    If click.Y < 0 Then
        click.Y = 0
    End If

    If click.Y > mnSourceHeight - (mnViewportHeight \ mlZoomFactor) Then
        click.Y = mnSourceHeight - (mnViewportHeight \ mlZoomFactor)
    End If

    Set BoundCheck = click

End Function

Public Property Let DisplayHeight(nNewValue As Integer)

    mnDisplayHeight = nNewValue

End Property

Public Property Let DisplayWidth(nNewValue As Integer)

    mnDisplayWidth = nNewValue

End Property

Public Property Get DisplayWidth() As Integer

    DisplayWidth = mnDisplayWidth

End Property

Public Property Get DisplayHeight() As Integer

    DisplayHeight = mnDisplayHeight

End Property

Public Property Let SourceHeight(nNewValue As Integer)

    mnSourceHeight = nNewValue

End Property

Public Property Let SourceWidth(nNewValue As Integer)

    mnSourceWidth = nNewValue

End Property

Public Property Get SourceWidth() As Integer

    SourceWidth = mnSourceWidth

End Property

Public Property Get SourceHeight() As Integer

    SourceHeight = mnSourceHeight

End Property

Public Property Get hScrollMax() As Integer

    hScrollMax = mnSourceWidth - (mnViewportWidth \ mlZoomFactor)
    
End Property

Public Property Get vScrollMax() As Integer

    vScrollMax = mnSourceHeight - (mnViewportHeight \ mlZoomFactor)

End Property

Public Function ScreenClickToGridClick(ByVal X As Integer, ByVal Y As Integer, GridWidth As Integer, GridHeight As Integer) As clsPoint
    
'***************************************************************************
'   Returns the value of the click taking the zoom factor into account
'***************************************************************************
    
    Dim click As New clsPoint
    Dim DispRatioX As Single
    Dim DispRatioY As Single
    
    DispRatioX = mnDisplayWidth / mnViewportWidth
    DispRatioY = mnDisplayHeight / mnViewportHeight
    
    click.X = Fix(((X + (mnViewportX * DispRatioX * mlZoomFactor)) / (GridWidth * DispRatioX)) / mlZoomFactor)
    click.Y = Fix(((Y + (mnViewportY * DispRatioY * mlZoomFactor)) / (GridHeight * DispRatioY)) / mlZoomFactor)
    
    Set ScreenClickToGridClick = click
    
End Function

Public Sub Draw(hDestDC As Long, Source As clsOffscreen)

'***************************************************************************
'   Draw information from a source offscreen to a DC
'***************************************************************************

    Source.BlitZoom hDestDC, 0, 0, CLng(mnDisplayWidth), CLng(mnDisplayHeight), CLng(mnViewportX), CLng(mnViewportY), mnViewportWidth \ mlZoomFactor, mnViewportHeight \ mlZoomFactor
    
End Sub
Public Property Get ViewportHeight() As Integer

    ViewportHeight = mnViewportHeight

End Property
Public Property Get ViewportWidth() As Integer

    ViewportWidth = mnViewportWidth

End Property

Public Property Let ViewportWidth(nNewValue As Integer)

    mnViewportWidth = nNewValue

End Property

Public Property Let ViewportHeight(nNewValue As Integer)

    mnViewportHeight = nNewValue

End Property

Public Property Let ViewportX(nNewValue As Integer)
    
    mnViewportX = nNewValue

End Property

Public Property Get ViewportX() As Integer

    ViewportX = mnViewportX

End Property

Public Property Let ViewportY(nNewValue As Integer)

    mnViewportY = nNewValue

End Property

Public Property Get ViewportY() As Integer

    ViewportY = mnViewportY

End Property

Public Sub ZoomIn()

    mlZoomFactor = mlZoomFactor * 2

End Sub
Public Sub ZoomOut()

    If mlZoomFactor = 1 Then
        Exit Sub
    End If
    
    mlZoomFactor = mlZoomFactor / 2

End Sub

Public Property Let Zoom(ZoomFactor As Integer)

    mlZoomFactor = ZoomFactor

End Property

Public Property Get Zoom() As Integer

    Zoom = mlZoomFactor

End Property

Public Sub ZoomInToClick(ByVal X As Integer, ByVal Y As Integer)

    ZoomIn
    
    Dim click As New clsPoint
    
    click.X = X - ((mnSourceWidth / mlZoomFactor) / 2)
    click.Y = Y - ((mnSourceHeight / mlZoomFactor) / 2)
    
    Set click = BoundCheck(click)
    
    mnViewportX = click.X
    mnViewportY = click.Y
    
End Sub


Public Sub ZoomOutToClick(ByVal X As Integer, ByVal Y As Integer)

    ZoomOut
    
    Dim click As New clsPoint
    
    click.X = X - ((mnSourceWidth / mlZoomFactor) / 2)
    click.Y = Y - ((mnSourceHeight / mlZoomFactor) / 2)
    
    Set click = BoundCheck(click)
    
    mnViewportX = click.X
    mnViewportY = click.Y
    
End Sub

Private Sub Class_Initialize()

    mlZoomFactor = 1

End Sub
