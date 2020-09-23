Attribute VB_Name = "mdlWindowProp"
' This program is made by a M. Schermer from the Netherlands
' I'm a prof. programmer for a lot of compagny's in the Netherlands and was in America
' a short time ago to program for a compagny that was making radars and that kind of stuff
' So you know from where you got this code

' Look in a couple of days on www.Planet-Source-Code.com for an update of this version.
' Sorry if there are some spelling faults in my notes or that there was left over some Dutch
' documentation

' Good luck with the code - Michiel.Schermer@Bit-ic.nl

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Public Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Hoort bij ShowWindow
Public Const SW_HIDE = 0 'Hide the window.
Public Const SW_MAXIMIZE = 3 'Maximize the window.
Public Const SW_MINIMIZE = 6 'Minimize the window.
Public Const SW_RESTORE = 9 'Restore the window (not maximized nor minimized).
Public Const SW_SHOW = 5 'Show the window.
Public Const SW_SHOWMAXIMIZED = 3 'Show the window maximized.
Public Const SW_SHOWMINIMIZED = 2 'Show the window minimized.
Public Const SW_SHOWMINNOACTIVE = 7 'Show the window minimized but do not activate it.
Public Const SW_SHOWNA = 8 'Show the window in its current state but do not activate it.
Public Const SW_SHOWNOACTIVATE = 4 'Show the window in its most recent size and position but do not activate it.
Public Const SW_SHOWNORMAL = 1 'Show the window and activate it (as usual).

'Hoort bij SendMessage
Public Const PBM_STEPIT = 1029
Public Const WM_USER = &H400
Public Const PBM_GETPOS = (WM_USER + 8)



