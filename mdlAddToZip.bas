Attribute VB_Name = "mdlAddToZip"
' This program is made by a M. Schermer from the Netherlands
' I'm a prof. programmer for a lot of compagny's in the Netherlands and was in America
' a short time ago to program for a compagny that was making radars and that kind of stuff
' So you know from where you got this code

' Look in a couple of days on www.Planet-Source-Code.com for an update of this version.
' Sorry if there are some spelling faults in my notes or that there was left over some Dutch
' documentation

' Good luck with the code - Michiel.Schermer@Bit-ic.nl

Public SetToHide As Boolean

Private Enum CheckDirOrFile
    IsDirectory = 1
    IsFilename = 2
End Enum


Public Function AddFilesToZip(strSourceDir As String, strDestFile As String, Optional LocationOfWinZip As String) As String
    Dim ProcessIsReady As Boolean
    Dim CheckFile As Boolean
    Dim CheckName As String
    Dim HwndOfWinzip As Integer
    Dim GetPercentDone As Byte
    
    CheckFile = FileExists(strDestFile, IsFilename)
    With Options
        If CheckFile = True Then
            Select Case .IfFileAlreadyExists
                Case 0: GoTo QuitZipProcess    '(Default) NotOverwrite
                Case 1: Kill strDestFile       'AlwaysOverwrite
                Case 2: GoTo QuitZipProcess    'NotOverwrite
                Case Else: GoTo QuitZipProcess '(Default) NotOverwrite
            End Select
        End If
    End With
    
    CheckFile = FileExists(strSourceDir, IsDirectory)
    If CheckFile = False Then
        MsgBox "Deze directory bestaat niet"
    Else
        Call RunWinZip(strSourceDir, strDestFile)
        CheckName = CheckNameToFind(strDestFile)
        HwndOfWinzip = FindHwndOfWinzip("WinZip - " & CheckName)
        
        Do Until ProcessIsReady = True
            GetPercentDone = GetPercentComplete(HwndOfWinzip)
            frmZipping.ProgressBar1.Value = GetPercentDone
            frmZipping.ProgressBar1.Refresh
            ProcessIsReady = CheckWinzipProcess
            DoEvents
        Loop
    End If
Exit Function

QuitZipProcess:
    MsgBox "Het ZIP process is afgebroken"
End Function

Private Function RunWinZip(strSourceDir As String, strDestFile As String, Optional LocationOfWinZip As String, Optional Options As String) As Integer
    Dim CommandLine As String
    
    CommandLine = GetOptions
    If Trim(LocationOfWinZip) = "" Then
        Shell "C:\program Files\Winzip\Winzip32.exe" & CommandLine & strDestFile & " " & strSourceDir & "\*.*", vbHide
    Else
        Shell LocationOfWinZip & CommandLine & strDestFile & " " & strSourceDir & "\*.*", vbHide
    End If
End Function

Private Function FindHwndOfWinzip(strNameToFind As String) As Integer
    Dim i As Integer
    Dim FindHWND As Integer
    
    Call DoEnumWindows

    For i = 1 To InsertWindowInfo.ItemsInArray
        FindHWND = InStr(LCase(InsertWindowInfo.GetWindowName(i)), LCase(strNameToFind))
        If FindHWND <> 0 Then
            GetWinZipHwnd = InsertWindowInfo.GetHWND(i)
            Call EnumChildWindows(GetWinZipHwnd, AddressOf WndEnumChildProc, Nothing)
            FindHwndOfWinzip = GetProgressHwnd
        End If
    Next i
End Function

Private Function GetPercentComplete(HwndToTrack As Integer) As Byte
    Dim CheckPercentComplete As Byte
    
    CheckPercentComplete = SendMessage(GetProgressHwnd, PBM_GETPOS, 0, ByVal 0)
        If CheckPercentComplete > 20 Then
            If SetToHide = False Then
                Call MoveWindowToHidePos
                SetToHide = True
                GetPercentComplete = CheckPercentComplete
                DoEvents
            Else
                GetPercentComplete = CheckPercentComplete
                DoEvents
            End If
        Else
            GetPercentComplete = CheckPercentComplete
            DoEvents
        End If
End Function

Private Sub MoveWindowToHidePos()
    Dim GetResX As Integer
    Dim GetResY As Integer

    GetResX = Screen.Width \ Screen.TwipsPerPixelX
    GetResY = Screen.Height \ Screen.TwipsPerPixelY

    Call MoveWindow(GetWinZipHwnd, GetResX + 1000, GetResY + 1000, 0, 0, 1)
    Call ShowWindow(GetWinZipHwnd, SW_SHOW)

End Sub

Private Function GetOptions() As String
    Dim MakeCmdLine As String
    
    With Options
        Select Case .ActionToDo
            Case 0, 1: MakeCmdLine = "-a"              'Add / -a
            Case 2: MakeCmdLine = "-f"                 'Freshen / -f
            Case 3: MakeCmdLine = "-u"                 'Update / u
            Case 4: MakeCmdLine = "-m"                 'Move / -m
            Case Else: MakeCmdLine = "-a"              'Add / -a
        End Select
        
        Select Case .Compression
            Case 0: MakeCmdLine = MakeCmdLine & " -en" '(Default) Normal / -en
            Case 1: MakeCmdLine = MakeCmdLine & " -ex" 'Extra / -ex
            Case 2: MakeCmdLine = MakeCmdLine & " -en" 'Normal / -en
            Case 3: MakeCmdLine = MakeCmdLine & " -ef" 'Fast / -ef
            Case 4: MakeCmdLine = MakeCmdLine & " -es" 'Super fast / -es
            Case 5: MakeCmdLine = MakeCmdLine & " -e0" 'No compression / -e0
            Case Else: MakeCmdLine = MakeCmdLine & " -en" '(Default) Normal / -en
        End Select
        
        Select Case .FilesToAdd
            Case 0, 1: MakeCmdLine = MakeCmdLine & " -hs" 'AddHiddenSystem / -hs
            Case Else: MakeCmdLine = MakeCmdLine & " -hs" '(Default) AddHiddenSystem / -hs
        End Select
        
        Select Case .Options
            Case 0: MakeCmdLine = MakeCmdLine & " -r"  '(Default) Recurse_Directories / -r
            Case 1: MakeCmdLine = MakeCmdLine & " -r"  'Recurse_Directories / -r
            Case 2: MakeCmdLine = MakeCmdLine & " -p"  'Save_Extra_Directory_Info / -p
            Case Else: MakeCmdLine = MakeCmdLine & " -r" '(Default) Recurse_Directories / -r
        End Select
        
        Select Case .PassWord
            Case 0:                                    'Nothing - No protection
            Case 1: MakeCmdLine = MakeCmdLine & " -s"  'Password protection / -s
            Case Else:                                 'Nothing - No protection
        End Select
    End With
    
    GetOptions = " " & MakeCmdLine & " "
End Function

Private Function CheckWinzipProcess() As Boolean
    Dim strNameOfWinzip As String
    
    strNameOfWinzip = GetWindowName(GetWinZipHwnd)
        If Trim(strNameOfWinzip) <> "" Then
            CheckWinzipProcess = False
            DoEvents
        Else
            CheckWinzipProcess = True
            DoEvents
        End If
End Function

Private Function FileExists(Filename, WhatToCheck As CheckDirOrFile) As Boolean
    Select Case WhatToCheck
        Case 1:
                If Right(Filename, 1) <> "\" Then
                    Filename = Filename & "\"
                    FileExists = (Dir(Filename) <> "")
                End If
        Case 2:
                FileExists = (Dir(Filename) <> "")
    End Select
End Function

Private Function CheckNameToFind(NameToCheck As String) As String
    If Mid(NameToCheck, 2, 2) = ":\" Then
        CheckNameToFind = Trim(Mid(NameToCheck, 4))
    End If
End Function

