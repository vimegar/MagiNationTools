Attribute VB_Name = "modGetPath"
Option Explicit
Public Function GetPathDialog() As String

    On Error Resume Next

    Load frmGetPath
    frmGetPath.Path = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
    frmGetPath.Show vbModal
    If frmGetPath.Cancelled Then
        Unload frmGetPath
        Exit Function
    End If
    GetPathDialog = frmGetPath.Path
    Unload frmGetPath

End Function


