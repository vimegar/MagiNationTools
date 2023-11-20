VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankPacker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BankPacker"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   Icon            =   "frmBankPacker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   Begin VB.CheckBox chkUsePrefix 
      Caption         =   "Use all s files"
      Height          =   495
      Left            =   7080
      TabIndex        =   50
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.FileListBox File 
      Height          =   2235
      Index           =   0
      Left            =   8640
      Pattern         =   "*.bin;vrm_*.s;spr_*.s;bg_*.s"
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Index           =   0
      Left            =   8640
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdPack 
      Caption         =   "&Pack"
      Height          =   360
      Left            =   7080
      TabIndex        =   34
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load..."
      Height          =   360
      Left            =   7080
      TabIndex        =   31
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "&Save..."
      Height          =   360
      Left            =   7080
      TabIndex        =   30
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "&Output..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   7080
      TabIndex        =   6
      Top             =   720
      Width           =   1200
   End
   Begin VB.Frame fraBanks 
      Caption         =   "B&anks"
      Height          =   4455
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.VScrollBar vsbBank 
         Height          =   3975
         Left            =   2640
         Max             =   118
         Min             =   1
         TabIndex        =   5
         Top             =   360
         Value           =   1
         Width           =   255
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   1
         Left            =   375
         TabIndex        =   8
         Top             =   705
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   2
         Left            =   375
         TabIndex        =   10
         Top             =   1065
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   3
         Left            =   375
         TabIndex        =   12
         Top             =   1425
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   4
         Left            =   375
         TabIndex        =   14
         Top             =   1785
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   5
         Left            =   375
         TabIndex        =   16
         Top             =   2145
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   6
         Left            =   375
         TabIndex        =   18
         Top             =   2505
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   7
         Left            =   375
         TabIndex        =   20
         Top             =   2865
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   8
         Left            =   375
         TabIndex        =   22
         Top             =   3225
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   9
         Left            =   375
         TabIndex        =   24
         Top             =   3585
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   10
         Left            =   375
         TabIndex        =   26
         Top             =   3945
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   38
         Top             =   345
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   10
         Left            =   1455
         TabIndex        =   49
         Top             =   3945
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   9
         Left            =   1455
         TabIndex        =   48
         Top             =   3585
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   8
         Left            =   1455
         TabIndex        =   47
         Top             =   3225
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   7
         Left            =   1455
         TabIndex        =   46
         Top             =   2865
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   6
         Left            =   1455
         TabIndex        =   45
         Top             =   2505
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   5
         Left            =   1455
         TabIndex        =   44
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   4
         Left            =   1455
         TabIndex        =   43
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   3
         Left            =   1455
         TabIndex        =   42
         Top             =   1425
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   2
         Left            =   1455
         TabIndex        =   41
         Top             =   1065
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   1
         Left            =   1455
         TabIndex        =   40
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label lblBankHex 
         AutoSize        =   -1  'True
         Caption         =   "&&H4000 (100%)"
         Height          =   195
         Index           =   0
         Left            =   1455
         TabIndex        =   39
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   225
      End
      Begin VB.Label lblBankNum 
         AutoSize        =   -1  'True
         Caption         =   "00:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Frame fraBinFiles 
      Caption         =   "&Bin Files"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdAddDir 
         Caption         =   "&Add Folder..."
         Height          =   360
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Width           =   1680
      End
      Begin VB.ComboBox cboUBound 
         Height          =   315
         ItemData        =   "frmBankPacker.frx":030A
         Left            =   1920
         List            =   "frmBankPacker.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3525
         Width           =   1695
      End
      Begin VB.ComboBox cboOrigin 
         Height          =   315
         ItemData        =   "frmBankPacker.frx":030E
         Left            =   120
         List            =   "frmBankPacker.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3525
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete File"
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   3960
         Width           =   1680
      End
      Begin VB.ListBox lstBinFiles 
         Height          =   2790
         ItemData        =   "frmBankPacker.frx":0312
         Left            =   120
         List            =   "frmBankPacker.frx":0314
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Upper Bound (In Hex):"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   33
         Top             =   3285
         Width           =   1590
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "&Origin (In Hex):"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   3285
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmBankPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_WIDTH = 8505
Private Const DEF_HEIGHT = 5085

Private Type tBank
    Size As Long
    FileCount As Integer
    Files() As String
    Paths() As String
End Type

Private mBPFile() As clsBPFile
Private mnBPFileCount As Integer

Private mtBanks() As tBank
Private mnStartBank As Integer
Private mnEndBank As Integer
Private mnRecCount As Integer
Private msRecPath As String
Private msOrgPath As String
Private mbErrors As Boolean
Private Sub mAddBPFile(Filename As String)
    
    Dim tempLen As Long
    Dim str As String
    Dim nFilenum As Integer
    
    nFilenum = FreeFile
 
    If InStr(UCase$(Filename), ".S") Then
        Open Filename For Input As #nFilenum
            Input #nFilenum, str
        Close #nFilenum
        
        tempLen = val(Mid$(str, 7))
            
        If tempLen = 0 Then
            nFilenum = FreeFile
            Open "c:\windows\desktop\bankpackererrors.txt" For Append As #nFilenum
                Print #nFilenum, Filename
            Close #nFilenum
            mbErrors = True
            Exit Sub
        End If
    End If
    
    mnBPFileCount = mnBPFileCount + 1
    ReDim Preserve mBPFile(mnBPFileCount)
    Set mBPFile(mnBPFileCount) = New clsBPFile

    With mBPFile(mnBPFileCount)
        .Filename = GetTruncFilename(Filename)
        
        If UCase$(Mid$(.Filename, Len(.Filename), 1)) = "S" Then
            .FileSize = tempLen
        Else
            .FileSize = FileLen(Filename)
        End If
        
        .FilePath = Mid$(Filename, InStr(Filename, msOrgPath), Len(Filename) - Len(.Filename) - InStr(Filename, msOrgPath))
    End With

End Sub


Private Sub mPackBanks()

    On Error GoTo HandleErrors

    Dim i As Integer
    Dim bigFile As Integer
    Dim curBank As Integer
    
    Screen.MousePointer = vbHourglass
    
    ReDim mtBanks(mnEndBank)
    
    For i = 0 To mnEndBank
        ReDim mtBanks(i).Files(mnBPFileCount)
        ReDim mtBanks(i).Paths(mnBPFileCount)
    Next i
    
    'quicksort array
    QuicksortBPFile mBPFile, 1, mnBPFileCount
    
    'pack banks
    For bigFile = mnBPFileCount To 1 Step -1
        
        'get curBank
        For i = mnStartBank To mnEndBank
            If mtBanks(i).Size + mBPFile(bigFile).FileSize <= &H4000 Then
                curBank = i
                Exit For
            End If
        Next i
    
        'pack
        mtBanks(curBank).FileCount = mtBanks(curBank).FileCount + 1
        mtBanks(curBank).Files(mtBanks(curBank).FileCount) = mBPFile(bigFile).Filename
        mtBanks(curBank).Paths(mtBanks(curBank).FileCount) = mBPFile(bigFile).FilePath
        mtBanks(curBank).Size = mtBanks(curBank).Size + mBPFile(bigFile).FileSize
    
    Next bigFile
    
    mUpdateBank

    Screen.MousePointer = 0

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmBankPacker:mPackBanks Error"
End Sub

Private Sub mUpdateBank()

    Dim i As Integer
    Dim str As String
    
    For i = 0 To 10
        
        'caption
        str = Hex(i + vsbBank.value - 1)
        If Len(str) < 2 Then
            str = "0" & str
        End If
        lblBankNum(i).Caption = str & ":"
        
        'progress bar
        If (mtBanks(i + vsbBank.value - 1).Size / &H4000) * 100 <= Progress(i).Max Then
            Progress(i).value = (mtBanks(i + vsbBank.value - 1).Size / &H4000) * 100
        Else
            Progress(i).value = Progress(i).Max
        End If
        
        'total hex thing
        lblBankHex(i).Caption = "&&H" & Format$(Hex(mtBanks(i + vsbBank.value - 1).Size), "0000") & " (" & Mid$(Format((Progress(i).value / Progress(i).Max), "percent"), 1, Len(Format((Progress(i).value / Progress(i).Max), "percent")) - 4) & "%" & ")"
        
    Next i

End Sub


Private Sub mUpdateList()

    lstBinFiles.Clear
    
    Dim i As Integer
    
    For i = 1 To mnBPFileCount
        lstBinFiles.AddItem GetRecentFilename(mBPFile(i).Filename) & " (H" & Hex(mBPFile(i).FileSize) & ")"
    Next i
    
    lstBinFiles.ListIndex = lstBinFiles.ListCount - 1
    
End Sub


Private Sub cboOrigin_Click()

    mnStartBank = val(HexToDec(cboOrigin.Text))
    If val(HexToDec(cboOrigin.Text)) + 1 <= vsbBank.Max Then
        vsbBank.value = val(HexToDec(cboOrigin.Text)) + 1
    Else
        vsbBank.value = vsbBank.Max
    End If
    
End Sub


Private Sub cboUBound_Click()

    mnEndBank = val(HexToDec(cboUBound.Text))

End Sub

Public Function AddBinFile(sPath As String) As Boolean

    Dim dirCursor As Integer
    Dim fileCursor As Integer
    
    If sPath = msRecPath Then
        Exit Function
    End If
    
    msRecPath = sPath
    
    mnRecCount = mnRecCount + 1
    On Error Resume Next
    Load Folder(mnRecCount)
    Load File(mnRecCount)
    On Error GoTo 0
    Folder(mnRecCount).Path = sPath
    
    If Folder(mnRecCount).ListCount > 0 Then
        Folder(mnRecCount).ListIndex = 0
    End If
    
    For dirCursor = Folder(mnRecCount).ListIndex To (Folder(mnRecCount).ListCount - 1)
        
        Dim bAdded As Boolean
        bAdded = AddBinFile(Folder(mnRecCount).List(dirCursor))
        
        File(mnRecCount).Path = Folder(mnRecCount).List(dirCursor)
        
        For fileCursor = 0 To (File(mnRecCount).ListCount - 1)
        
            DoEvents
        
            If Not bAdded Then
                mAddBPFile Folder(mnRecCount).List(dirCursor) & "\" & File(mnRecCount).List(fileCursor)
            End If
            
            AddBinFile = True
            
        Next fileCursor
        
    Next dirCursor
    
    On Error Resume Next
    Unload Folder(mnRecCount)
    Unload File(mnRecCount)
    On Error GoTo 0
    mnRecCount = mnRecCount - 1
    
End Function

Private Sub cmdAddDir_Click()

    Dim sPath As String
    sPath = GetPathDialog
    
    Screen.MousePointer = vbHourglass
    
    msRecPath = ""
    mnRecCount = 0
    
    Dim i As Integer
    Dim str As String
    msOrgPath = sPath
    For i = Len(msOrgPath) To 1 Step -1
        If Mid$(msOrgPath, i, 1) = "\" Then
            str = Mid$(msOrgPath, i)
            Exit For
        End If
    Next i
    msOrgPath = str
    
    DeleteFile "c:\windows\desktop\bankpackererrors.txt"
    
    If chkUsePrefix.value = vbUnchecked Then
        File(0).Pattern = "*.bin;vrm_*.s;spr_*.s;bg_*.s"
    Else
        File(0).Pattern = "*.bin;*.s"
    End If
    
    AddBinFile sPath
    Screen.MousePointer = 0
    If mbErrors Then
        mbErrors = False
        MsgBox "There were one or more '.s' files that did not contain the necessary header information to be included in this project.  Please refer to 'c:\windows\desktop\bankpackererrors.txt' for the filenames.", vbInformation, "Input Error"
    End If
    
    mUpdateList

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


Private Sub cmdDelete_Click()

    If lstBinFiles.ListIndex < 0 Then
        Exit Sub
    End If
    
    Dim startIndex As Integer
    Dim i As Integer
    
    startIndex = lstBinFiles.ListIndex
    
    For i = lstBinFiles.ListIndex + 1 To mnBPFileCount - 1
        mBPFile(i) = mBPFile(i + 1)
    Next i
    
    ReDim Preserve mBPFile(UBound(mBPFile) - 1)
    
    mUpdateList
    
    If lstBinFiles.ListCount = 0 Then
        cmdOutput.Enabled = False
    End If
    
    If startIndex > (lstBinFiles.ListCount - 1) Then
        lstBinFiles.ListIndex = (lstBinFiles.ListCount - 1)
    ElseIf lstBinFiles.ListCount = 0 Then
        lstBinFiles.ListIndex = -1
    Else
        lstBinFiles.ListIndex = startIndex
    End If

End Sub



Private Sub cmdLoad_Click()

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Load BankPacker Project"
        .Filename = ""
        .Filter = "BankPacker Project (*.bpp)|*.bpp"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename

        Dim i As Integer
        Dim d As Integer
        Dim nFilenum As Integer
        nFilenum = FreeFile

        Open .Filename For Input As #nFilenum
        
        Input #nFilenum, d
        cboOrigin.ListIndex = d
        
        Input #nFilenum, d
        cboUBound.ListIndex = d
        
        Input #nFilenum, d
        
        ReDim mBPFile(d)
        
        Dim s As String
        
        For i = 1 To UBound(mBPFile)
            
            Input #nFilenum, s
            mBPFile(i).Filename = s
            
            Input #nFilenum, s
            mBPFile(i).FilePath = s
            
            Input #nFilenum, d
            mBPFile(i).FileSize = d
            
        Next i
        
        Close #nFilenum

    End With

    mUpdateList
    mPackBanks

End Sub

Private Sub cmdOutput_Click()
    
    On Error GoTo HandleErrors
    
    Dim sPath As String
    Dim sFilename As String
    Dim i As Integer
    Dim j As Integer
    Dim strTab As String
    Dim strBankNum As String
    Dim strDPath As String
    Dim dummy As String
    
    Screen.MousePointer = vbHourglass
    
    sPath = GetPathDialog
    If sPath = "" Then GoTo oEnd
    
    strTab = Chr(vbKeyTab)
    
    Dim nFilenum As String
    nFilenum = FreeFile
    
    For i = mnStartBank To mnEndBank
        
        strBankNum = Hex(i)
        If Len(strBankNum) < 2 Then
            strBankNum = "0" & strBankNum
        End If
        
        Dim str As String
        str = Format$(Hex(i \ 8), "00")
        
        If Len(str) < 2 Then
            str = "0" & str
        End If
        
        CreateDir sPath & "\MBit" & str
        sFilename = sPath & "\MBit" & str & "\Bank" & strBankNum & ".s"
        
        Open sFilename For Output As #nFilenum
    
            Print #nFilenum, ";********************************"
            Print #nFilenum, "; " & UCase$(GetTruncFilename(sFilename))
            Print #nFilenum, ";********************************"
            Print #nFilenum, ";" & strTab & "Author:" & strTab & "Mr. Yuk"
            Print #nFilenum, ";" & strTab & "(c)2000" & strTab & "Interactive Imagination"
            Print #nFilenum, ";" & strTab & "All rights reserved"
            Print #nFilenum, ""
            Print #nFilenum, ";********************************"
    
            Print #nFilenum, "BANK" & strBankNum & strTab & strTab & "GROUP" & strTab & "$" & strBankNum
            Print #nFilenum, strTab & strTab & strTab & "ORG" & strTab & strTab & "$4000"
            Print #nFilenum, ""
            
            For j = 1 To mtBanks(i).FileCount
                dummy = UCase$(Mid$(GetTruncFilename(mtBanks(i).Files(j)), 1, Len(GetTruncFilename(mtBanks(i).Files(j))) - 4))
                
                If Len(mtBanks(i).Paths(j)) > 0 Then
                    If Mid$(mtBanks(i).Paths(j), Len(mtBanks(i).Paths(j)) - 1) = "\" Then
                        strDPath = ""
                    Else
                        strDPath = "\"
                    End If
                Else
                    strDPath = ""
                End If
                
                If InStr(UCase$(mtBanks(i).Files(j)), ".S") = 0 Then
                    Print #nFilenum, dummy & String((8 - Len(dummy)) * -(Len(dummy) <= 8), " ") & strTab & "LIBBIN" & strTab & Mid$(UCase$(mtBanks(i).Paths(j)), 2) & strDPath & UCase$(GetTruncFilename(mtBanks(i).Files(j)))
                Else
                    Print #nFilenum, strTab & strTab & strTab & strTab & "LIB" & strTab & strTab & Mid$(UCase$(mtBanks(i).Paths(j)), 2) & strDPath & UCase$(GetTruncFilename(mtBanks(i).Files(j)))
                End If
            Next j
    
            Print #nFilenum, ""
            Print #nFilenum, ";********************************"
            Print #nFilenum, strTab & "END"
            Print #nFilenum, ";********************************"
        
        Close #nFilenum
    Next i
    
oEnd:
    Screen.MousePointer = 0
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmBankPacker:cmdOutput_Click Error"
    Screen.MousePointer = 0
End Sub
Private Sub cmdPack_Click()
    
    mPackBanks
    cmdOutput.Enabled = True

End Sub

Private Sub cmdSaveAs_Click()

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save BankPacker Project"
        .Filename = ""
        .Filter = "BankPacker Project (*.bpp)|*.bpp"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename

        Dim i As Integer
        Dim nFilenum As Integer
        nFilenum = FreeFile

        Open .Filename For Output As #nFilenum
        
        Write #nFilenum, mnStartBank
        Write #nFilenum, mnEndBank
        
        Write #nFilenum, UBound(mBPFile)
        
        For i = 1 To UBound(mBPFile)
            Write #nFilenum, mBPFile(i).Filename
            Write #nFilenum, mBPFile(i).FilePath
            Write #nFilenum, mBPFile(i).FileSize
        Next i
        
        Close #nFilenum

    End With

End Sub
Private Sub Form_Load()

    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT

    ReDim mBPFile(0)
    mnBPFileCount = 0
    Set mBPFile(0) = New clsBPFile
    ReDim mtBanks(128)

    Dim i As Integer
    Dim str As String
    
    cboOrigin.Clear
    
    For i = 0 To 127
        str = Hex(i)
        If Len(str) < 2 Then
            str = "0" & str
        End If
        cboOrigin.AddItem str
        cboUBound.AddItem str
    Next i
    
    cboOrigin.ListIndex = 0
    cboUBound.ListIndex = 127

    mUpdateBank
    
End Sub







Private Sub vsbBank_Change()

    mUpdateBank

End Sub


Private Sub vsbBank_Scroll()

    vsbBank_Change
    
End Sub


