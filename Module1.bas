Attribute VB_Name = "Module1"
Public Declare Function EnumChildWindows Lib "User32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public GetProgressHwnd As Long
    Public GetWinZipHwnd As Long
    
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
