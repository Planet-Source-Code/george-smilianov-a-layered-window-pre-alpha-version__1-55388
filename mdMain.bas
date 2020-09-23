Attribute VB_Name = "mdMain"
'|---------------------------------------------------------------------------------------------------------------
'| A "layered" window
'| Version 0.1 / 5th Aug 2004
'|
'| Copyright (c) 2004 George Smilianov
'| www.smilianov.net; smilianov@dir.bg; ICQ UIN: 173377653
'|
'| If you like this code, please vote for me at Planet Source Code
'|-----------------------------------------------------------------------------------------------------------------
'|-----------------------------------------------------------------------------------------------------------------


Option Explicit

Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
    Public Const GW_CHILD As Long = 5
    Public Const GW_HWNDLAST As Long = 1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Const MK_LBUTTON = &H1
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202
    
    Public Const MK_RBUTTON = &H2
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205

    Public Const WM_LBUTTONDBLCLK = &H203

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Const GWL_EXSTYLE = -20
    Public Const WS_EX_LAYERED = &H80000
    Public Const GWL_STYLE = (-16)
    Public Const WS_VISIBLE = &H10000000
    
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Const LWA_ALPHA = &H2
    
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Type POINTAPI
        x As Long
        y As Long
    End Type
    
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

