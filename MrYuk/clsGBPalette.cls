VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGBPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'   VB Interface Setup
'***************************************************************************

    Option Explicit
    Implements intResource
    Implements intResourceClient

'***************************************************************************
'   Class properties
'***************************************************************************

    Private mColors(32) As New clsRGB
    Private miStart As Byte
    Private miEnd As Byte
    Private msParentPath As String

'***************************************************************************
'   Resource cache object pointers
'***************************************************************************

    Private mResourceCacheObject As clsResourceCacheObject

Public Property Let intResource_ParentPath(RHS As String)

    msParentPath = RHS

End Property

Public Property Get intResource_ParentPath() As String

    intResource_ParentPath = msParentPath

End Property

Public Property Get CacheObject() As clsResourceCacheObject

    Set CacheObject = intResource_CacheObject

End Property


Public Sub GetPalFromBMP(BMPOffscreen As clsOffscreen)

    Dim i As Long
    Dim oRGB As clsRGB
    
    For i = 0 To 31
        Set oRGB = GetRGBFromLong(BMPOffscreen.GetPixel(i, 0))
        mColors(i + 1).Red = oRGB.Red \ 8
        mColors(i + 1).Green = oRGB.Green \ 8
        mColors(i + 1).Blue = oRGB.Blue \ 8
    Next i

End Sub

Public Property Let iEnd(iNewValue As Byte)

    miEnd = iNewValue

End Property

Public Property Get iEnd() As Byte

    iEnd = miEnd

End Property

Public Property Let iStart(iNewValue As Byte)

    miStart = iNewValue

End Property

Public Property Get iStart() As Byte

    iStart = miStart

End Property

Public Sub UpdateClients(iClient As intResourceClient)

    intResource_UpdateClients iClient

End Sub

Public Sub PackToBin(ByVal sFilename As String)
    
'    BinFile     binFile;
'    u8          pal;
'    u8          color;
'    u16         pack;
'    u8          low;
'    u8          high;

'    if ( !BinFile_Create ( &binFile, fileName )) return 0;

'    for ( pal = 0; pal < CGBPAL_PALS; pal++ )
'        for ( color = 0; color < CGBPAL_COLORS; color++ )
'            {
'            pack = ((cgbPal->Colors[pal][color].Red>>3)<<CGBPAL_REDSHIFT);
'            pack += ((cgbPal->Colors[pal][color].Green>>3)<<CGBPAL_GREENSHIFT);
'            pack += ((cgbPal->Colors[pal][color].Blue>>3)<<CGBPAL_BLUESHIFT);

'            low = (u8)(pack & 0xFF);
'            high = (u8)(pack>>8);

'            BinFile_Write ( &binFile, &low, sizeof ( u8 ));
'            BinFile_Write ( &binFile, &high, sizeof ( u8 ));
'            }

'    BinFile_Close ( &binFile );
'    return 1;
    
    If sFilename = "" Then
        Exit Sub
    End If
    
    Dim nFilenum As Integer
    nFilenum = FreeFile
    
    Open sFilename For Binary Access Write As #nFilenum
    
    Dim pal As Byte
    Dim color As Byte
    Dim Pack As Integer
    Dim mung As Byte
    Dim gumbo As Byte
    Dim Index As Byte
    Dim i As Integer
          
    For pal = iStart To iEnd
        For color = 0 To 3
           
            Index = color + (pal * 4) + 1
            
            Pack = mColors(Index).Red
            Pack = Pack + mColors(Index).Green * 32
            Pack = Pack + mColors(Index).Blue * 1024
            
            mung = CByte(Pack And &HFF)
            gumbo = CByte(Pack \ 256)
            
            Put #nFilenum, , mung
            Put #nFilenum, , gumbo
            
        Next color
    Next pal
        
    Close #nFilenum

End Sub
Public Property Set Colors(Index As Integer, oNewValue As clsRGB)

    Set mColors(Index) = oNewValue

End Property

Public Property Get Colors(Index As Integer) As clsRGB

    Set Colors = mColors(Index)

End Property

Private Sub Class_Initialize()

'***************************************************************************
'   Initialize the color values to white (RGB 255, 255, 255)
'***************************************************************************

    Dim i As Integer
    
'    mColorTable(0!).Red = 32
'    mColorTable(1!).Red = 44
'    mColorTable(2!).Red = 56
'    mColorTable(3!).Red = 68
'    mColorTable(4!).Red = 80
'    mColorTable(5!).Red = 92
'    mColorTable(6!).Red = 104
'    mColorTable(7!).Red = 116
'    mColorTable(8!).Red = 124
'    mColorTable(9!).Red = 136
'    mColorTable(10).Red = 144
'    mColorTable(11).Red = 156
'    mColorTable(12).Red = 164
'    mColorTable(13).Red = 172
'    mColorTable(14).Red = 180
'    mColorTable(15).Red = 184
'    mColorTable(16).Red = 192
'    mColorTable(17).Red = 196
'    mColorTable(18).Red = 200
'    mColorTable(19).Red = 204
'    mColorTable(20).Red = 208
'    mColorTable(21).Red = 212
'    mColorTable(22).Red = 212
'    mColorTable(23).Red = 216
'    mColorTable(24).Red = 220
'    mColorTable(25).Red = 224
'    mColorTable(26).Red = 224
'    mColorTable(27).Red = 228
'    mColorTable(28).Red = 228
'    mColorTable(29).Red = 228
'    mColorTable(30).Red = 228
'    mColorTable(31).Red = 228

'    mColorTable(0!).Green = 32
'    mColorTable(1!).Green = 44
'    mColorTable(2!).Green = 56
'    mColorTable(3!).Green = 68
'    mColorTable(4!).Green = 84
'    mColorTable(5!).Green = 96
'    mColorTable(6!).Green = 108
'    mColorTable(7!).Green = 120
'    mColorTable(8!).Green = 132
'    mColorTable(9!).Green = 140
'    mColorTable(10).Green = 152
'    mColorTable(11).Green = 160
'    mColorTable(12).Green = 172
'    mColorTable(13).Green = 180
'    mColorTable(14).Green = 188
'    mColorTable(15).Green = 192
'    mColorTable(16).Green = 200
'    mColorTable(17).Green = 204
'    mColorTable(18).Green = 208
'    mColorTable(19).Green = 212
'    mColorTable(20).Green = 216
'    mColorTable(21).Green = 220
'    mColorTable(22).Green = 224
'    mColorTable(23).Green = 228
'    mColorTable(24).Green = 228
'    mColorTable(25).Green = 232
'    mColorTable(26).Green = 236
'    mColorTable(27).Green = 236
'    mColorTable(28).Green = 240
'    mColorTable(29).Green = 240
'    mColorTable(30).Green = 240
'    mColorTable(31).Green = 240
    
'    mColorTable(0!).Blue = 32
'    mColorTable(1!).Blue = 44
'    mColorTable(2!).Blue = 56
'    mColorTable(3!).Blue = 68
'    mColorTable(4!).Blue = 84
'    mColorTable(5!).Blue = 96
'    mColorTable(6!).Blue = 108
'    mColorTable(7!).Blue = 120
'    mColorTable(8!).Blue = 132
'    mColorTable(9!).Blue = 140
'    mColorTable(10).Blue = 152
'    mColorTable(11).Blue = 160
'    mColorTable(12).Blue = 172
'    mColorTable(13).Blue = 180
'    mColorTable(14).Blue = 188
'    mColorTable(15).Blue = 192
'    mColorTable(16).Blue = 200
'    mColorTable(17).Blue = 204
'    mColorTable(18).Blue = 208
'    mColorTable(19).Blue = 212
'    mColorTable(20).Blue = 216
'    mColorTable(21).Blue = 220
'    mColorTable(22).Blue = 224
'    mColorTable(23).Blue = 228
'    mColorTable(24).Blue = 228
'    mColorTable(25).Blue = 232
'    mColorTable(26).Blue = 236
'    mColorTable(27).Blue = 236
'    mColorTable(28).Blue = 240
'    mColorTable(29).Blue = 240
'    mColorTable(30).Blue = 240
'    mColorTable(31).Blue = 240
    
End Sub

Private Property Get intResource_CacheObject() As clsResourceCacheObject

    Set intResource_CacheObject = mResourceCacheObject

End Property

Private Sub intResource_Delete()

'***************************************************************************
'   Class properties
'***************************************************************************

    'Dim i As Integer
    
    'For i = 0 To 32
    '    Set mColors(i) = Nothing
    'Next i

End Sub

Public Sub Delete()

    intResource_Delete

End Sub

Private Function intResource_Unpack(nFilenum As Integer) As Boolean
    
'***************************************************************************
'   Unpack a palette file from a .pal
'***************************************************************************
    
    On Error GoTo HandleErrors
    
    intResource_Unpack = True
    
    Dim i As Integer
    Dim dummyR As Byte
    Dim dummyG As Byte
    Dim dummyB As Byte
    
    For i = 1 To 32
        
        DoEvents
        
        Get #nFilenum, , dummyR
        mColors(i).Red = CByte(dummyR And &H1F)
      
        Get #nFilenum, , dummyG
        mColors(i).Green = CByte(dummyG And &H1F)
      
        Get #nFilenum, , dummyB
        mColors(i).Blue = CByte(dummyB And &H1F)
        
    Next i
    
    Dim iByte As Byte
    Get #nFilenum, , iByte
    
    If iByte > 0 Then
        Get #nFilenum, , miStart
        Get #nFilenum, , miEnd
    Else
        miStart = 0
        miEnd = 7
    End If
    
Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "clsGBPalette:intResource_Unpack Error"
    intResource_Unpack = False
End Function

Private Sub intResource_Pack(nFilenum As Integer)

'***************************************************************************
'   Pack to a .pal file
'***************************************************************************

    Put #nFilenum, , CByte(GB_PALETTE)
    
    Dim i As Integer
    Dim dummy As Byte
    
    For i = 1 To 32
    
        dummy = CByte(mColors(i).Red And &H1F)
        Put #nFilenum, , dummy
    
        dummy = CByte(mColors(i).Green And &H1F)
        Put #nFilenum, , dummy
    
        dummy = CByte(mColors(i).Blue And &H1F)
        Put #nFilenum, , dummy
    
    Next i

    Put #nFilenum, , CByte(1)
    Put #nFilenum, , miStart
    Put #nFilenum, , miEnd

End Sub

Private Property Set intResource_CacheObject(oNewValue As clsResourceCacheObject)
    
    Set mResourceCacheObject = oNewValue
    
End Property

Private Sub intResource_UpdateClients(iClient As intResourceClient)
    
'***************************************************************************
'   Update all clients of this resource
'***************************************************************************
    
    If Not mResourceCacheObject Is Nothing Then
        mResourceCacheObject.UpdateClients iClient
    End If
    
End Sub

Private Property Get intResource_ResourceType() As GB_FILETYPES
    
'***************************************************************************
'   Return the resource type
'***************************************************************************
    
    intResource_ResourceType = GB_PALETTE

End Property

Private Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

End Sub


