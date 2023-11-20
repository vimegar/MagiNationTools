Attribute VB_Name = "modBMP"
Option Explicit

Public Enum BPP_TYPES
'these types are not available yet
    'BPP1 = 1
    'BPP4 = 4
    'BPP8 = 8
    'BPP16 = 16
    BPP24 = 24
End Enum
Public Sub SaveToBMP(DestFilename As String, HBitmapDC As Long, BitsPerPixel As BPP_TYPES, BitmapWidth As Long, BitmapHeight As Long, Optional Progress As ProgressBar)

    On Error GoTo HandleErrors

    DeleteFile DestFilename
    
'Create windows bitmap structures
    Dim bitSize As Long
    
    bitSize = ((BitmapWidth * BitmapHeight) * 3)
    
    'BITMAP FILE HEADER : Size = 14 bytes
    Dim bmfh As BITMAPFILEHEADER
    With bmfh
        .bfType = 19778 'BM
        .bfSize = 14 + 40 + bitSize
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = .bfSize - bitSize
    End With
    
    'BITMAP INFO HEADER : Size = 40 bytes
    Dim bmih As BITMAPINFOHEADER
    With bmih
        .biSize = 40
        .biWidth = BitmapWidth
        .biHeight = BitmapHeight
        .biPlanes = 1
        .biBitCount = BitsPerPixel
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    
    Dim X As Long
    Dim Y As Long
    Dim lRGB As Long
    Dim nFilenum As Integer
    
'Output bitmap file
    nFilenum = FreeFile

    Open DestFilename For Binary As #nFilenum
    
        Put #nFilenum, , bmfh
        Put #nFilenum, , bmih
        
        If Not Progress Is Nothing Then
            Progress.Max = (BitmapWidth * BitmapHeight)
            Progress.value = 0
        End If
        
        For Y = (BitmapHeight - 1) To 0 Step -1
            
            DoEvents
            
            If gbAbort Then
                GoTo Abort
            End If
            
            For X = 0 To (BitmapWidth - 1) Step 1
            
                lRGB = GetPixel(HBitmapDC, X, Y)

                Put #nFilenum, , CByte(Fix(lRGB / 65536) And &HFF)
                Put #nFilenum, , CByte(Fix(lRGB / 256) And &HFF)
                Put #nFilenum, , CByte(lRGB And &HFF)

                If Not Progress Is Nothing Then
                    Progress.value = Progress.value + 1
                End If

            Next X
        Next Y

    
rExit:
    Progress.value = 0
    Close #nFilenum
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "modBMP:SaveToBMP Error"
    GoTo rExit

Abort:
    Progress.value = 0
    Close #nFilenum
    DeleteFile DestFilename
End Sub


