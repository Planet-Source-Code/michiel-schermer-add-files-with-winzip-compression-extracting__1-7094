Attribute VB_Name = "mdlEnumWindows"
' This program is made by a M. Schermer from the Netherlands
' I'm a prof. programmer for a lot of compagny's in the Netherlands and was in America
' a short time ago to program for a compagny that was making radars and that kind of stuff
' So you know from where you got this code

' Look in a couple of days on www.Planet-Source-Code.com for an update of this version.
' Sorry if there are some spelling faults in my notes or that there was left over some Dutch
' documentation

' Good luck with the code - Michiel.Schermer@Bit-ic.nl

Private Declare Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "User32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
    Public GetProgressHwnd As Long
    Public GetWinZipHwnd As Long
    
    Public InsertWindowInfo As GetWindowInfo
        Private Type GetWindowInfo
            GetHWND(1 To 1000) As Integer
            GetWindowName(1 To 1000) As String
            ItemsInArray As Integer
        End Type

Public Sub DoEnumWindows()
    Call EnumWindows(AddressOf EnumWindowProc, &H0)
End Sub

Private Function EnumWindowProc(ByVal hwnd As Long, ToListbox As ListBox) As Long
    Dim strWindowName    As String
    Dim strClassName     As String
    
    strWindowName = GetWindowName(hwnd)
    
        With InsertWindowInfo
            .ItemsInArray = .ItemsInArray + 1
            .GetHWND(.ItemsInArray) = CInt(hwnd)
            .GetWindowName(.ItemsInArray) = strWindowName
        End With
    
    EnumWindowProc = 1 'Zorgt ervoor dat Enumwindows door blijft gaan totdat er geen HWND's meer zijn
End Function

Public Function GetWindowName(Handle As Long) As String
    Dim intWindowLenght  As Integer
    Dim strWindowName    As String
        
        intWindowLenght = GetWindowTextLength(Handle) + 1
        strWindowName = Space$(intWindowLenght)
        GetWindowText Handle, strWindowName, intWindowLenght ' API function call
        strWindowName = Mid(strWindowName, 1, Len(strWindowName) - 1)
    
    GetWindowName = strWindowName
End Function
    
Public Function WndEnumChildProc(ByVal hwnd As Long) As Long
    Dim bRet As Long
    Dim myStr As String * 50
    Dim FindClass As Integer
    
    bRet = GetClassName(hwnd, myStr, 50)
    
    FindClass = InStr(LCase(myStr), "msctls_progress32")
        If FindClass <> 0 Then
            GetProgressHwnd = hwnd
        Else
            WndEnumChildProc = 1
        End If
End Function
