VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim hBGWindow As Long
Dim pt As POINTAPI
Dim nMousePosition As Long

Public Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Private Sub Form_DblClick()
    GetCursorPos pt
    nMousePosition = MakeDWord(CInt(pt.x), CInt(pt.y))
    PostMessage hBGWindow, WM_LBUTTONDBLCLK, MK_LBUTTON, nMousePosition
End Sub

Private Sub Form_Load()
    hBGWindow = GetWindow(GetWindow(GetWindow(GetWindow(GetDesktopWindow(), GW_CHILD), GW_HWNDLAST), GW_CHILD), GW_CHILD)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 100, LWA_ALPHA
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    pt.x = ScaleX(x, ScaleMode, vbPixels)
    pt.y = ScaleY(y, ScaleMode, vbPixels)
    Me.Caption = pt.x & "," & pt.y & " " & hBGWindow
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GetCursorPos pt
    nMousePosition = MakeDWord(CInt(pt.x), CInt(pt.y))
    
    If Button = 1 Then
        PostMessage hBGWindow, WM_LBUTTONDOWN, MK_LBUTTON, nMousePosition
        PostMessage hBGWindow, WM_LBUTTONUP, MK_LBUTTON, nMousePosition
    Else
        PostMessage hBGWindow, WM_RBUTTONDOWN, MK_RBUTTON, nMousePosition
        PostMessage hBGWindow, WM_RBUTTONUP, MK_RBUTTON, nMousePosition
    End If
End Sub
