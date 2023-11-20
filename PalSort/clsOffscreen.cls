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

    Rect 0, 0, mlWidth, mlHeight, vbWhite
    
End Sub

Public Sub LineDraw(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, lColor As Long, Optional Rect As Boolean)
    
    Dim hPen As Long
    Dim hPenOld As Long
    Dim hBrush As Long
    Dim hBrushOld As Long
    Dim point As POINTAPI
    Dim linX1 As Integer
    Dim linY1 As Integer
    Dim linX2 As Integer
    Dim linY2 As Integer
    
    If Rect Then
    
        linX1 = x1: linY1 = y1: linX2 = x2: linY2 = y1
        GoSub Lin
        
        linX1 = x1: linY1 = y1: linX2 = x1: linY2 = y2
        GoSub Lin
        
        linX1 = x1: linY1 = y2: linX2 = x2 + 1: linY2 = y2
        GoSub Lin
        
        linX1 = x2: linY1 = y1: linX2 = x2: linY2 = y2 + 1
        GoSub Lin
        
    Else
        
        linX1 = x1: linY1 = y1: linX2 = x2: linY2 = y2
        GoSub Lin
        
    End If


Exit Sub

Lin:
    MoveToEx hDC, linX1, linY1, point
    
    hPen = CreatePen(PS_SOLID, 1, lColor)
    hBrush = CreateSolidBrush(lColor)
        
    hPenOld = SelectObject(hDC, hPen)
    hBrushOld = SelectObject(hDC, hBrush)
    
    LineTo hDC, linX2, linY2

    SelectObject hDC, hPenOld
    DeleteObject hPen
    
    SelectObject hDC, hBrushOld
    DeleteObject hBrush
    
    Return
End Sub

Public Sub Blit(hDestDC As Long, XOffset As Long, yOffset As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC
'***************************************************************************
    
    BitBlt hDestDC, XOffset, yOffset, mlWidth, mlHeight, hDC, 0, 0, SRCCOPY
    
End Sub

Public Sub BlitRaster(hDestDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, RasterOp As RasterOpConstants)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC using raster operations
'***************************************************************************
    Dim d As Integer
    
    d = BitBlt(hDestDC, x, y, nWidth, nHeight, mlHDC, xSrc, ySrc, RasterOp)
    
End Sub

Public Sub BlitRect(hDestDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC in a rectangle
'***************************************************************************
    
    BitBlt hDestDC, x, y, nWidth, nHeight, mlHDC, xSrc, ySrc, SRCCOPY
    
End Sub

Public Sub BlitZoom(hDestDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, xSrc As Long, ySrc As Long, srcWidth As Long, srcHeight As Long)
    
'***************************************************************************
'   Transfer pixels from this offscreen surface to a target DC with an aspect ratio
'***************************************************************************
    
    StretchBlt hDestDC, x, y, nWidth, nHeight, hDC, xSrc, ySrc, srcWidth, srcHeight, SRCCOPY
    
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
    mlHDC = CreateCompatibleDC(frmPalSort.hDC)
    mlHBitmap = CreateCompatibleBitmap(frmPalSort.hDC, mlWidth, mlHeight)
    SelectObject mlHDC, mlHBitmap
    Rectangle mlHDC, -1, -1, mlWidth + 1, mlHeight + 1
    
End Sub

Public Sub CreateBitmapFromBMP(ByVal sFilename As String)

'***************************************************************************
'   Create a GB bitmap from a windows source bitmap
'***************************************************************************

    Dim pic As Picture
    Dim hImageDC As Long
    Dim bm As Bitmap
    Dim hBM As Long
    
'Load windows bitmap
    Set pic = LoadPicture(sFilename)
    hBM = pic.Handle
    gdi32.GetObject hBM, Len(bm), bm
    
'Blit the windows bitmap onto the offscreen surface
    hImageDC = CreateCompatibleDC(frmPalSort.hDC)
    SelectObject hImageDC, hBM
    
    Create bm.bmWidth + 1, bm.bmHeight + 1
    
    BitBlt hDC, 0, 0, mlWidth, mlHeight, hImageDC, 0, 0, SRCCOPY
    
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

Public Sub SetPixel(x As Long, y As Long, color As Long)
    
'***************************************************************************
'   Change the color of target pixel to a specified color
'***************************************************************************
    
    gdi32.SetPixel mlHDC, x, y, color
    
End Sub

Public Sub Rect(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)

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
    
    Rectangle mlHDC, x1, y1, x2, y2

    SelectObject mlHDC, hPenOld
    DeleteObject hPen
    SelectObject mlHDC, hBrushOld
    DeleteObject hBrush

End Sub

Public Function GetPixel(x As Long, y As Long) As Long

'***************************************************************************
'   Read the color value of target pixel
'***************************************************************************

    GetPixel = gdi32.GetPixel(mlHDC, x, y)

End Function

Public Property Get Width() As Long
    
    Width = mlWidth
    
End Property

Public Property Let Width(lNewValue As Long)
    
    mlWidth = lNewValue
    
End Property

Public Property Get Height() As Long
    
    Height = mlHeight
    
End Property

Public Property Let Height(lNewValue As Long)
    
    mlHeight = lNewValue
    
End Property

Public Property Get hDC() As Long

    hDC = mlHDC

End Property

Private Sub Class_Terminate()

    Delete
    
End Sub


