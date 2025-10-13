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



Private Sub Form_Load()
    'Call SysLoad00_SysStartLoad(tglDockingTrgtFrm)

End Sub

Private Sub btnSysStart_Click()
'Stop
    Call SysLoad00_SysStartLoad(tglDockingTrgtFrm)
    
End Sub


Private Sub btnSair_Click()

'Stop
    Call Scr_ShowAccess
    DoCmd.Close
    
End Sub


Private Sub bTest_Click()

    Dim vA, vB, vC, vD, vE, vF
    Dim vCmb1 As Variant, vLst1 As Variant, vTxt1 As Variant, vGrp1 As Variant, vChk1 As Variant
    Dim vCmb2 As Variant, vLst2 As Variant, vTxt2 As Variant, vGrp2 As Variant, vChk2 As Variant
    Dim cCmb As Control, cLst As Control, cTxt As Control, cGrp As Control, cChk As Control

MsgBox bTest.ControlType

    'cCtrL = IIf(cCtrL.DefaultValue = "", Null, Replace(cCtrL.DefaultValue, """", ""))
    'Call HighlightClrChange(Int(cCtrL.ControlType), cCtrL, True)


    Set cGrp = Me.grpSrcCor
    Set cChk = Me.chkSaldo
    Set cCmb = Me.cmbSrcCateg
    Set cLst = Me.lstProdsBase
    Set cTxt = Me.txtSrchMed
Stop
    vGrp1 = Me.grpSrcCor.Value
    vChk1 = Me.chkSaldo.Value
    vCmb1 = Me.cmbSrcCateg.Value
    vLst1 = Me.lstProdsBase.Value
    vTxt1 = Me.txtSrchMed.Value
Stop
    cGrp = 3
    cChk = False
    cCmb = 1
    cLst = 1
    cTxt = "texto novo"
    
    
'    cgrp =
'    chk =
'    ccmb=
'    clst=
'    ctxt =
    
'    Me.grpSrcCor.Value = Me.grpSrcCor.DefaultValue
'    Me.chkSaldo.Value = Me.chkSaldo.DefaultValue
'    Me.cmbSrcCateg.Value = Me.cmbSrcCateg.DefaultValue
'    Me.lstProdsBase.Value = Me.lstProdsBase.DefaultValue
'
'
'    Me.grpSrcCor.Value = Me.grpSrcCor.DefaultValue
'    Me.chkSaldo.Value = Me.chkSaldo.DefaultValue
'    Me.cmbSrcCateg.Value = Me.cmbSrcCateg.DefaultValue
'    Me.lstProdsBase.Value = Me.lstProdsBase.DefaultValue
    
    
Stop

End Sub

Private Sub btnCheckColor_Click()
    Dim vA, vB, vC
    Dim cTrgtCtrl As Control
    Dim sStR1 As String, sStR2 As String
    Dim sTrgtCtrl As String
    'Dim cColorValuesLabel As Control
'Stop
    
'    On Error Resume Next
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
    
    'Atribuição dos valores a [ clObjFormOpenParams ] usados nas rotinas de abertura de formulários
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams
        
        'Garante que a variável [ bTrgtFormIsDocked ] que indica se o [ TrgtForm ] deve abrir ancorado
        ' receba False caso não tenha sido pressionada
        If Not IsNull(tglDockingTrgtFrm.Value) Then
            Set .cTglBtnDocking = tglDockingTrgtFrm
            If Not .cTglBtnDocking Is Nothing Then .bFrmIsDocked = .cTglBtnDocking.Value
        ElseIf ControlExists("tglDockingTrgtFrm", Me) Then Set .cTglBtnDocking = tglDockingTrgtFrm
        End If

        '.bIsTrggForm = False
        .bSetPosition = True
        '.fTrgtForm   este valor deve ser definido após a abertura do formulário em FormLoad05_OpenForm
        .lngFormLeft = GbLngTrgtLeft
        .lngFormTop = GbLngTrgtTop
        '.lngRecID = 30
        .sFormMode = "ProdView"
        .sTrggForm = Me.Name
        .sTrgtForm = "frm_01(1)cProdEstoque"
    
    End With

    'Inicia as rotinas pra abertura de formulário
    Call FormLoad01_FormLoadingStart(clObjFormOpenParams)
    
    'Abre o form [ frm_00(1)bSysTimer ] para monitorar "FormMove" tanto do TrggForm quanto do TrgtForm
    DoCmd.OpenForm "frm_00(1)bSysTimer", acNormal, , , , acHidden
    
End Sub
