Attribute VB_Name = "Módulo 10a - pbSubs (Ctrls manipulat)"
Option Compare Database
Option Explicit


'Código pra rolar textboxes
Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
Private Const sModName = "Module1_scroll"
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const SB_PAGEUP = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_LINEUP = 0
Public Const SB_LINEDOWN = 1
Public Declare PtrSafe Function trans_msg Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

'Código pra rolar textboxes
Function Fun_handler(CNTL As Control) As Long
    CNTL.SetFocus
    Fun_handler = apiGetFocus

End Function
