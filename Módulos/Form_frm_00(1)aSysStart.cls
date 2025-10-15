VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_00(1)aSysStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cmbSrcCateg_BeforeUpdate(Cancel As Integer)
    Dim vA, vB, vC, vD, vE
    Dim cCtrL As Control
    
    Set cCtrL = cmbSrcCateg
Stop
    
    vB = cCtrL.ListIndex


End Sub

Private Sub Form_Load()
    
'    Call SysLoad00_SysStartLoad

End Sub

Private Sub btnSysStart_Click()
'Stop
    Call SysLoad00_SysStartLoad
    
End Sub


Private Sub btnSair_Click()

'Stop
    Call Scr_ShowAccess
    DoCmd.Close
    
End Sub


Private Sub bTest_Click()
   Dim vA, vB, vC, vD, vE
   Dim cCtrL1 As Control, cCtrL2 As Control
   
   Set cCtrL1 = Me.lstTeste1
   Set cCtrL2 = Me.lstTeste2
   

Stop
   vA = cCtrL1.Value
   'vB = cCtrL1.DefaultValue
    
    
    vC = cCtrL1.ListIndex
    vD = cCtrL1.ItemsSelected.Count
    'vE = cCtrL1.ItemsSelected.Item(0)
    
    
   vA = cCtrL2.Value
   'vB = cCtrL2.DefaultValue
    
    
    vC = cCtrL2.ListIndex
    vD = cCtrL2.ItemsSelected.Count
    'vE = cCtrL2.ItemsSelected.Item(0)
    
    
    
End Sub

Private Sub btnCheckColor_Click()
    Dim vA, vB, vC
    Dim cTrgtCtrl As Control
    Dim sStR1 As String, sStR2 As String
    Dim sTrgtCtrl As String
    'Dim cColorValuesLabel As Control
'Stop
    
    Set cTrgtCtrl = Screen.PreviousControl
    
    'Se tiver havido o erro [ 2483 ] indicando que não foi selecionado nenhum controle pra testar a cor
    ' [ sTrgtCtrL ] ficará vazio provocando a saída da rotina
    sTrgtCtrl = cTrgtCtrl.Name
        
    
    If sTrgtCtrl = "" Or sTrgtCtrl = "btnSair" Then
    
        sStR1 = "Selecione um controle pra testar a cor" & vbCr & "-------------------------------------------------------------------------------"
        
        Call msgboxErrorAlert(sStR1, sStR2)
        Exit Sub
    
    End If
    
    TempVars("sTrgtCtrL") = cTrgtCtrl.Name
    TempVars("sForM") = Me.Name
'Stop

    'Me.lblColorCheck.Caption = "Color Check   [ " & cTrgtCtrl.Name & " ]"
    DoCmd.OpenForm "frmDev_90dDevPickColor", acNormal
    
    'Set cColorValuesLabel = Me.txtBoxColorValue
    'Call ColorCheck(cTrgtCtrl, Me.txtBoxColorValue)  '- versão antiga do Color Picker
    
    
End Sub

Private Sub lstProdsBase_AfterUpdate()
    Dim vA, vB, vC, vD, vE
    Dim cCtrL As Control
    
    Set cCtrL = lstProdsBase
Stop
    vA = cCtrL.ItemsSelected.Count
    vB = cCtrL.ListIndex
    
    vD = cCtrL.Value
    'vE = cCtrL.Column(0, 3)
    'vA = cCtrL.Column(0, vB)
    
    
Stop
    vA = cCtrL.ListIndex
    vB = cCtrL.ItemsSelected.Count
    
    cCtrL.Selected(1) = True
    vB = cCtrL.ItemsSelected.Count
    
    For Each vA In cCtrL.ItemsSelected
        vC = cCtrL.Column(0, vA)
        
    Next vA
    
    
    
    
'
'
'    vA = cCtrL.ListIndex
'    vB = cCtrL.ItemsSelected.Count
'
'    cCtrL.Selected(vA) = True
'    vB = cCtrL.ItemsSelected.Count
'

End Sub

Private Sub Lista246_AfterUpdate()
    Dim vA, vB
Stop
    vA = Lista246.ItemsSelected.Count
    
    Lista246.Selected(Lista246.ListIndex) = True
    
End Sub



'Private Sub btnInitEvents_Click()
'    Dim vA, vB, vC
'    Dim cCtrL As Control
'    Dim sForM As String
'
'    sForM = Me.Name
'    Set cCtrL = Me.txtSrchMed
'
'
'Stop
'    If dictCtrlsEvents.Exists(cCtrL.Name) = True Then  'verifica se o controle já foi armazenado
'        Set clObjCtrlsEvents = dictCtrlsEvents(cCtrL.Name)
'
'    Else
'        Set clObjCtrlsEvents = New cls_10aCtrls_Events
'        dictCtrlsEvents.Add cCtrL.Name, clObjCtrlsEvents
'
'
'    End If
'
'    Set clObjCtrlsEvents.InitCtrl = cCtrL
'
'
'End Sub


Private Sub tglDockingTrgtFrm_Click()
    
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams
        Set .cTglBtnDocking = Me.tglDockingTrgtFrm
        Set .fTrgtForm = Me
        .sTrgtForm = Me.Name
        .bFrmIsDocked = .cTglBtnDocking.Value
        
    End With
    
    Call FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams)
    
End Sub


Private Sub btnProdEstoque_Click()
    Dim vA, vB

' muda o Sort da lista


    
End Sub
