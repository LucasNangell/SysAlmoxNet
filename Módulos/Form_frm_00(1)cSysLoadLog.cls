VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_00(1)cSysLoadLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCopyText_Click()
    Dim vA, vB
    Dim cCtrL As Control
'Stop
    'Seleciona todo o texto de [ txtLoadLog ] pra permitir
    ' que o texto seja copiado
    Set cCtrL = Me.txtLoadLog
    cCtrL.SetFocus
    cCtrL.SelStart = 0
    cCtrL.SelLength = Len(cCtrL.Text)
    
End Sub

Private Sub btnSair_Click()
'    DoCmd.Close
    
End Sub

'Private Sub btnSair_Click()
'
'    DoCmd.Close
'
'End Sub


Private Sub Form_Load()
    
'Stop
    If clObjFormOpenParams.vOpenArgs <> "" Then
        Me.lblTrggForm.Caption = "[ " & clObjFormOpenParams.sTrggForm & " ]"
        Me.txtLoadLog = clObjFormOpenParams.vOpenArgs
                
        Call TextboxScrollWhenNeeded(Me.txtLoadLog)
    End If
'Stop
    Set clObjFormOpenParams.fTrgtForm = Me
    
    On Error GoTo -1
    Call FormLoad06a_BackFromFormLoad(clObjFormOpenParams)
    On Error GoTo -1
    
'Stop

End Sub
Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal nline As Long) 'ByVal Page As Boolean, ByVal Count As Long)
    
'    'código antigo
'    If txtLoadLog.ScrollBars Then txtLoadLog.SetFocus
'    txtLoadLog.SelStart = IIf(txtLoadLog.SelStart + Count * 25 < 0, 0, txtLoadLog.SelStart + Count * 25)
    
    'código novo
    'Dim vA, vB
    Dim Hcntrl As Long
    Dim n As Long
    Dim fActualForm As Form
    
    Set fActualForm = Me
'Stop
    'vA = fActualForm.ActiveControl.Name
'    Debug.Print vA
    
    'rotina pra rolar textboxes de múltiplas linhas pela Roda do Mouse
    If fActualForm.ActiveControl.ControlType = acTextBox Then
    'If ActiveControl.Properties.Item("controltype") = acTextBox Then
        Hcntrl = Fun_handler(fActualForm.ActiveControl)
        'Hcntrl = Fun_handler(Screen.ActiveControl)
        For n = 1 To Abs(nline)
            trans_msg Hcntrl, WM_VSCROLL, IIf(nline < 0, SB_LINEUP, SB_LINEDOWN), 0&

        Next

    End If
    

End Sub

Private Sub lblStatusBar_DblClick(Cancel As Integer)
'Stop
    Call FormStatusBar04_OpnLogForm(Me)
    
End Sub
