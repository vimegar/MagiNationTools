VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOffscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlHBitmap As Long
Private mlHDC As Long
Private mlWidth As Long
Private mlHeight As Long

Public Sub Cls()

    RECT 0, 0, mlWidth, mlHeight, vbWhite
    
End Sub

Public Sub LineDraw(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, lColor As Long, Optional RECT As Boolean)
    
    Dim hPen As Long
    Dim hPenOld As Long
    Dim hBrush As Long
    Dim hBrushOld As Long
    Dim point As POINTAPI
    Dim linX1 As Integer
    Dim linY1 As Integer
    Dim linX2 As Integer
    Dim linY2 As Integer
    
    If RECT Then
    
        linX1 = X1: linY1 = Y1: linX2 = X2: linY2 = Y1
        GoSub Lin
        
        linX1 = X1: linY1 = Y1: linX2 = X1: linY2 = Y2
        GoSub Lin
        
        linX1 = X1: linY1 = Y2: linX2 = X2 + 1: linY2 = Y2
        GoSub Lin
        
        linX1 = X2: linY1 = Y1: linX2 = X2: linY2 = Y2 + 1
        GoSub Lin
        
    Else
        
        linX1 = X1: linY1 = Y1: linX2 = X2: linY2 = Y2
        GoSub Lin
        
    End If


Exit Sub

Lin:
    MoveToEx hdc, linX1, linY1, point
    
    hPen = CreatePen(PS_SOLID, 1, lColor)
    hBrush = CreateSolidBrush(lColor)
        
    hPenOld = SelectObject(hdc, hPen)
    hBrushOld = SelectObject(hdc, hBrush)
    
    LineTo hdc, linX2, linY2

    SelectObject hdc, hPenOld
    DeleteObject hPen
    
    SelectObject hdc, hBrushOld
    DeleteObject hBrush
    
    Return
End Sub

Public Sub Blit(hDestDC As Long, XOffset As Long, yOffset As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC
'***************************************************************************
    
    BitBlt hDestDC, XOffset, yOffset, mlWidth, mlHeight, hdc, 0, 0, vbSrcCopy
    
End Sub

Public Sub BlitRaster(hDestDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, RasterOp As RasterOpConstants)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC using raster operations
'***************************************************************************
    Dim d As Integer
    
    d = BitBlt(hDestDC, X, Y, nWidth, nHeight, mlHDC, xSrc, ySrc, RasterOp)
    
End Sub

Public Sub BlitRect(hDestDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC in a rectangle
'***************************************************************************
    
    BitBlt hDestDC, X, Y, nWidth, nHeight, mlHDC, xSrc, ySrc, vbSrcCopy
    
End Sub

Public Sub BlitZoom(hDestDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, srcWidth As Long, srcHeight As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC with an aspect ratio
'***************************************************************************
    
    StretchBlt hDestDC, X, Y, nWidth, nHeight, hdc, xSrc, ySrc, srcWidth, srcHeight, vbSrcCopy
    
End Sub


Public Sub Create(nWidth As Long, nHeight As Long)
    
'***************************************************************************
'   Initialize the offscreen surface
'***************************************************************************
    
    If mlHBitmap <> 0 Then
        Delete
    End If
    
    mlWidth = nWidth
    mlHeight = nHeight
    
    DoEvents
'Create blank offscreen object
    mlHDC = CreateCompatibleDC(glMainHDC)
    mlHBitmap = CreateCompatibleBitmap(glMainHDC, mlWidth, mlHeight)
    SelectObject mlHDC, mlHBitmap
    Rectangle mlHDC, -1, -1, mlWidth + 1, mlHeight + 1
    
End Sub

Public Sub CreateBitmapFromBMP(ByVal sFilename As String)

'***************************************************************************
'   Create a GB bitmap from a windows source bitmap
'***************************************************************************

    Dim hImageDC As Long
    Dim bm As BITMAP
    Dim hBM As Long
    Dim pic As Picture
    
'Load windows bitmap
    Set pic = LoadPicture(sFilename)
    hBM = pic.Handle
    modWinAPI.GetObject hBM, Len(bm), bm
    
'Blit the wi0ndows bitmap onto the offscreen surface
    hImageDC = CreateCompatibleDC(glMainHDC)
    SelectObject hImageDC, hBM
    
    Create bm.bmWidth + 1, bm.bmHeight + 1
    
    BitBlt hdc, 0, 0, mlWidth, mlHeight, hImageDC, 0, 0, vbSrcCopy
    
    DeleteDC hImageDC
    DeleteObject hBM
    
End Sub


Public Sub Delete()
    
'***************************************************************************
'   Release memory
'***************************************************************************
    
    DeleteObject mlHBitmap
    DeleteDC mlHDC
    
End Sub

Public Sub SetPixel(X As Long, Y As Long, color As Long)
    
'***************************************************************************
'   Change the color of target pixel to a specified color
'***************************************************************************
    
    modWinAPI.SetPixel mlHDC, X, Y, color
    
End Sub

Public Sub RECT(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

'***************************************************************************
'   Draw a rectangle on the target surface using a specified color
'***************************************************************************

    Dim hPen As Long
    Dim hPenOld As Long
    Dim hBrush As Long
    Dim hBrushOld As Long
    
    hPen = CreatePen(PS_SOLID, 2, lColor)
    hBrush = CreateSolidBrush(lColor)
        
    hPenOld = SelectObject(mlHDC, hPen)
    hBrushOld = SelectObject(mlHDC, hBrush)
    
    Rectangle mlHDC, X1, Y1, X2, Y2

    SelectObject mlHDC, hPenOld
    DeleteObject hPen
    SelectObject mlHDC, hBrushOld
    DeleteObject hBrush

End Sub

Public Function GetPixel(X As Long, Y As Long) As Long

'***************************************************************************
'   Read the color value of target pixel
'***************************************************************************

    GetPixel = modWinAPI.GetPixel(mlHDC, X, Y)

End Function

Public Property Get width() As Long
    
    width = mlWidth
    
End Property

Public Property Let width(lNewValue As Long)
    
    mlWidth = lNewValue
    
End Property

Public Property Get height() As Long
    
    height = mlHeight
    
End Property

Public Property Let height(lNewValue As Long)
    
    mlHeight = lNewValue
    
End Property

Public Property Get hdc() As Long

    hdc = mlHDC

End Property

Private Sub Class_Terminate()

    Delete
    
End Sub


