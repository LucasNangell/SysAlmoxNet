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
    'Call pbSub22_TargtCtrlsQryFieldsBuild(Me.Name, lstProdsBase, "01")
    Call SysLoad00_SysStartLoad

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
   Dim vA, vB, vC
   Dim sTitulo As String, sTexto As String, sTrggForM As String
   sTrggForM = Me.Name

'Stop
    sTitulo = "Cancelamento de Ediçao de Dados"
    sTexto = "As alterações realizadas serão perdidas." & vbCrLf & "Confirma?"
    vA = DialogBoxReply(sTitulo, sTexto, sTrggForM)

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
    Dim vA
Stop
    vA = lstProdsBase.ItemsSelected.Count
    

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
        .sTrggForM = Me.Name
        .sTrgtForm = "frm_01(1)cProdEstoque"
    
    End With

    'Inicia as rotinas pra abertura de formulário
    Call FormLoad01_FormLoadingStart(clObjFormOpenParams)
    
    'Abre o form [ frm_00(1)bSysTimer ] para monitorar "FormMove" tanto do TrggForm quanto do TrgtForm
    DoCmd.OpenForm "frm_00(1)bSysTimer", acNormal, , , , acHidden
    
End Sub
