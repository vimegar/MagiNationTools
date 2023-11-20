Attribute VB_Name = "modMain"
Option Explicit

'Windows API declares
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public gHero As New clsHero
Public gCreatures(3) As New clsCreature

Public Function GetIniData(ApplicationName As String, KeyName As String, IniFilename As String) As String
'********************************************
'
'   ApplicationName     Identification string
'   KeyName             Tag of the data to be retrieved
'   IniFilename         String containing the filename of
'                       the .ini file
'
'   This function returns the data stored within the ini
'   file that is specified.
'
'********************************************
        
    Dim dummy As String * 32
    
    GetPrivateProfileString ApplicationName, KeyName, "", dummy, 32, IniFilename
    GetIniData = dummy
    
End Function


Public Sub WriteIniData(ApplicationName As String, KeyName As String, IniFilename As String, DataString As String)
'********************************************
'
'   ApplicationName     Identification string
'   KeyName             Tag of the data to be retrieved
'   IniFilename         String containing the filename of
'                       the .ini file
'   DataString          The string to be written to the ini.
'
'   This procedure saves DataString under the tag KeyName
'   located in the ini file IniFilename.
'
'********************************************
        
    WritePrivateProfileString ApplicationName, KeyName, DataString, IniFilename
    
End Sub
