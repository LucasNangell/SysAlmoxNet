Attribute VB_Name = "Módulo 00e - Aux (Access window)"
Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long ) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
'Constantes pra manipulação da janela do access
Public Const SW_HIDE As Long = 0
Public Const SW_SHOW As Long = 5
Public Const SW_RESTORE As Long = 9
Public Const SW_MINIMIZE As Long = 6

'Constantes pra manipulação do formulário
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_SHOWWINDOW As Long = &H40


'Função para deixar um formulário no topo da pilha
Public Sub Scr_FormAlwaysOnTop(frm As Access.Form, Optional StayOnTop As Boolean = True)
    Dim hWndForm As LongPtr
    
    hWndForm = frm.hWnd
    
    SetWindowPos hWndForm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    If Not StayOnTop Then SetWindowPos hWndForm, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
  
  End Sub


Public Sub Scr_HideAccess()
    ShowWindow Application.hWndAccessApp, SW_HIDE
    
End Sub


Public Sub Scr_ShowAccess()
    ShowWindow Application.hWndAccessApp, SW_SHOW

End Sub


'Public Sub Scr_RestoreAccess()
'    ShowWindow Application.hWndAccessApp, SW_RESTORE
'
'End Sub
