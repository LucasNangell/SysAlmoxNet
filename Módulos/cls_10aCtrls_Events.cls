VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_10aCtrls_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public sCtrlName As String  'Nome do controle

Private WithEvents mLabel As Label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mToggleButton As ToggleButton
Attribute mToggleButton.VB_VarHelpID = -1
Private WithEvents mCommandButton As CommandButton
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mCheckBox As CheckBox
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mListBox As ListBox
Attribute mListBox.VB_VarHelpID = -1
Private WithEvents mOptGroup As OptionGroup
Attribute mOptGroup.VB_VarHelpID = -1
Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1


'------------------------------------------------------------
'Trecho pra manipular eventos dos controles
'----------------------------------
Public Property Set InitCtrl(cCtrl As Control)
    Dim vA, vB
    Dim sCtrL As String

    'vA = bOnDirty
    sCtrL = cCtrl.Name
    vB = cCtrl.ControlType

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "InitCtrl: [ " & sCtrl & " ]"
'Stop

'    mCheckBox.AfterUpdate
'    mComboBox.AfterUpdate
'    -- mComboBox.OnChange -- >   mComboBox.AfterUpdate
'    mListBox.AfterUpdate
'    mOptGroup.AfterUpdate
'    mTextBox.AfterUpdate
'    mTextBox.OnChange
    If gBbEnableErrorHandler Then On Error Resume Next
'Stop
    Select Case cCtrl.ControlType
        Case acCheckBox
            'Set mCheckBox = Nothing
            Set mCheckBox = cCtrl
            cCtrl.AfterUpdate = "[Event Procedure]"
            GoTo CaseResume     'por algum motivo quando o controle é um CheckBox o código
                                ' é direcionado pro Case Else ao invés de executar o acCheckBox e sair do Case
                                ' por esse motivo é necessário forçar sua saída
        
        Case acComboBox
            Set mComboBox = Nothing
            Set mComboBox = cCtrl
'Stop
            cCtrl.OnChange = "[Event Procedure]"
            cCtrl.AfterUpdate = "[Event Procedure]"

        Case acCommandButton
            Set mCommandButton = Nothing
            Set mCommandButton = cCtrl
            
            cCtrl.OnClick = "[Event Procedure]"

        Case acLabel
            Set mLabel = Nothing
            Set mLabel = cCtrl
            
            cCtrl.OnDblClick = "[Event Procedure]"

        Case acListBox
            Set mListBox = Nothing
            Set mListBox = cCtrl
            cCtrl.AfterUpdate = "[Event Procedure]"
            cCtrl.OnLostFocus = "[Event Procedure]"
            cCtrl.BeforeUpdate = "[Event Procedure]"
        Case acOptionGroup
'Stop
            Set mOptGroup = Nothing
            Set mOptGroup = cCtrl
            cCtrl.AfterUpdate = "[Event Procedure]"
            'GoTo CaseResume

        Case acTextBox
            Set mTextBox = Nothing
            Set mTextBox = cCtrl
'Stop
            cCtrl.OnChange = "[Event Procedure]"
            cCtrl.AfterUpdate = "[Event Procedure]"
            cCtrl.OnGotFocus = "[Event Procedure]"
            'cCtrl.BeforeUpdate = "[Event Procedure]"
            'cCtrl.OnKeyPress = "[Event Procedure]"
            cCtrl.OnLostFocus = "[Event Procedure]"
            
        Case acToggleButton
            Set mToggleButton = Nothing
            Set mToggleButton = cCtrl
            
            cCtrl.OnClick = "[Event Procedure]"
            
        Case Else
'Stop
            Exit Property
    
    End Select
'Stop
CaseResume:
'Stop
End Property


Private Sub mComboBox_AfterUpdate()
    Dim fForm As Form
    Dim cCtrl As Control
    
    Set cCtrl = mComboBox
    Set fForm = cCtrl.Parent
    
    'Aciona a atualização do [ TargetCtrl ] associado ao controle
    Call pb_TargtCtrlUpdate00_TimerDelay(fForm, cCtrl)
    

End Sub


Private Sub mCommandButton_Click()
    Dim vA
    Dim sMode As String, sForM As String, sActType As String, sFilGrp As String
    Dim cCtrl As Control, cEdtCtrl As Control
    Dim sCtrL As String
    Dim fForm As Form
    Dim sLstbxTag As String
    Dim vSplittedTAG As Variant
    Dim vKeyCtrl As Variant
    
'MsgBox "testando RstArea commandbutton event"
'Stop
    On Error Resume Next
    Set cCtrl = mCommandButton
    sCtrL = cCtrl.Name
    Set fForm = cCtrl.Parent
    
    If Err.Number = 5 Then Exit Sub
    On Error GoTo -1
    
    If gBbEnableErrorHandler Then On Error Resume Next

    
    If Err.Number = 5 Then Exit Sub
    On Error GoTo -1

    sForM = cCtrl.Parent.Name
    
    Set clObjCommButtons = dictFormCommButtons(sForM)(cCtrl.Name)

    'Nos botões tipo [ RstArea ] o sActType retornado é "RstArea0#"
    ' o Select Case foi modificado pra permitir usar a função like
    ' e verificar se o início da string tem o termo "RstArea"
    sActType = clObjCommButtons.sActType
    
    Select Case True
        
        Case sActType Like "CancelEdit"
'Stop
            If TempVars("bReply") = True Then
                sFilGrp = clObjCommButtons.sFilGrp
                If sFilGrp <> "" Then
                    Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
                    sCtrL = clObjTargtCtrlParam.sTargtCtrlName
                    
                    If ControlExists(sCtrL, fForm) Then
                        Set cCtrl = fForm.Controls(sCtrL)
                        If fForm.Controls(sCtrL).ControlType = acListBox Then Call PbSubFillFieldsByList(cCtrl)
                        On Error GoTo -1
    'Stop
                    End If
                End If
                
                On Error GoTo -1
                Call pbSub00_CtrlsEnblDsble_GetParams(cCtrl.Parent, clObjCommButtons.sFrmMode, sFilGrp)
                If gBbEnableErrorHandler Then On Error GoTo -1: On Error Resume Next
                
                'Mantém botão de edição habilitado caso haja algum item selecionado na lista
                If cCtrl.ItemsSelected.Count > 0 Then
                    
                    Call pbSub00_CtrlsEnblDsble_GetParams(fForm, "RecSelect", sFilGrp)
                End If
            End If
            TempVars.RemoveAll
        Case sActType Like "CancelNew"
            sFilGrp = clObjCommButtons.sFilGrp
            If sFilGrp <> "" Then
                For Each vKeyCtrl In dictFormDataFlds00Ctrls(sForM)
                    Set clObjCtrlDataFieds = dictFormDataFlds00Ctrls(sForM)(vKeyCtrl)
                    If clObjCtrlDataFieds.sDataFieldGrp = sFilGrp Then
                        If ControlExists(sCtrL, fForm) Then fForm.Controls(sCtrL).Value = ""
                    End If
                Next vKeyCtrl
            End If
            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrl.Parent, clObjCommButtons.sFrmMode, sFilGrp)
        
        Case sActType Like "CloseForm"
            DoCmd.Close
                      
        Case sActType Like "EditRec"
            sFilGrp = clObjCommButtons.sFilGrp
            If sFilGrp <> "" Then
                Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
                sCtrL = clObjTargtCtrlParam.sTargtCtrlName
    
                If ControlExists(sCtrL, fForm) Then
                    Set cCtrl = fForm.Controls(sCtrL)
                    If cCtrl.ControlType = acListBox Then
                    
                        If cCtrl.ListIndex > -1 Then
                            
                            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrl.Parent, clObjCommButtons.sFrmMode, sFilGrp)
                            For Each vKeyCtrl In dictFormCommButtons(sForM)
                                Set clObjCommButtons = dictFormCommButtons(sForM)(vKeyCtrl)
                                If clObjCommButtons.sFilGrp = sFilGrp Then
                                    If InStr(clObjCommButtons.sActType, "Cancel") > 0 Then clObjCommButtons.sActType = "CancelEdit"
                                    If InStr(clObjCommButtons.sActType, "Save") > 0 Then clObjCommButtons.sActType = "SaveEdit"
                                End If
                            Next vKeyCtrl
                        Else
                            MsgBox "Selecione qual registro você deseja editar clicando na lista.", vbInformation
                        End If
                    End If
                    
                End If
            End If
                      
        Case sActType Like "HideAcc"
            'Oculta a janela do Access
            Call Scr_HideAccess
            
            'Traz o formulário pra frente pra reexibi-lo
            Call Scr_FormAlwaysOnTop(cCtrl.Parent, False)

        Case sActType Like "NewRec"
            
            sFilGrp = clObjCommButtons.sFilGrp
            If sFilGrp <> "" Then
                For Each vKeyCtrl In dictFormDataFlds00Ctrls(sForM)
                    Set clObjCtrlDataFieds = dictFormDataFlds00Ctrls(sForM)(vKeyCtrl)
                    If clObjCtrlDataFieds.sDataFieldGrp = sFilGrp Then
                        sCtrL = vKeyCtrl
                        If ControlExists(sCtrL, fForm) Then fForm.Controls(sCtrL).Value = ""
                    End If
                Next vKeyCtrl
                
                For Each vKeyCtrl In dictFormCommButtons(sForM)
                    Set clObjCommButtons = dictFormCommButtons(sForM)(vKeyCtrl)
                    If clObjCommButtons.sFilGrp = sFilGrp And InStr(clObjCommButtons.sActType, "Save") > 0 Then clObjCommButtons.sActType = "SaveNew"
                Next vKeyCtrl
            End If
            
            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrl.Parent, clObjCommButtons.sFrmMode, sFilGrp)


        Case sActType Like "OpenForm"
            
            sFilGrp = clObjCommButtons.sFilGrp
            Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
            sCtrL = clObjTargtCtrlParam.sTargtCtrlName
            Set cCtrl = fForm.Controls(sCtrL)
            
            'Atribuição dos valores a [ clObjFormOpenParams ] usados nas rotinas de abertura de formulários
            Set clObjFormOpenParams = New cls_09cParamsToOpenForms
            
            With clObjFormOpenParams
                
                'Garante que a variável [ bTrgtFormIsDocked ] que indica se o [ TrgtForm ] deve abrir ancorado
                ' receba False caso não tenha sido pressionada
                If Not IsNull(cCtrl.Parent.tglDockingTrgtFrm.Value) Then
                    Set .cTglBtnDocking = cCtrl.Parent.tglDockingTrgtFrm
                    If Not .cTglBtnDocking Is Nothing Then .bFrmIsDocked = .cTglBtnDocking.Value
                ElseIf ControlExists("tglDockingTrgtFrm", cCtrl.Parent) Then Set .cTglBtnDocking = cCtrl.Parent.tglDockingTrgtFrm
                End If
                .sFilGrp = sFilGrp
                .sFormMode = clObjCommButtons.sFrmMode
                .sTrggForM = cCtrl.Parent.Name
                .sTrgtForm = clObjCommButtons.sForM
                .lngRecID = cCtrl.ListIndex
                .bShowRecID = True
                .bSetPosition = True
                .bCentralizeForm = True
                Set .cTrgtCtrl = dictFormFilterGrpTrgts(cCtrl.Parent.Name)(sFilGrp)
                .sQryFieldID = GetsQryFieldID(clObjTargtCtrlParam)
            End With

                        
            'Inicia as rotinas pra abertura de formulário
            Call FormLoad01_FormLoadingStart(clObjFormOpenParams)
            
            'Abre o form [ frm_00(1)bSysTimer ] para monitorar "FormMove" tanto do TrggForm quanto do TrgtForm
            DoCmd.OpenForm "frm_00(1)bSysTimer", acNormal, , , , acHidden

        Case sActType Like "RstArea*"   '* é usado como Wildcard
'Stop
            Call pb_TargtCtrlUpdate04_RESETarea(cCtrl.Parent.Name, cCtrl.Name)

            
        Case sActType Like "Save*"      '* é usado como Wildcard
            sFilGrp = clObjCommButtons.sFilGrp
            Call PbSubRecDataFields(cCtrl)
            For Each vA In dictFormFilterGrpTrgts(sForM)(sFilGrp)
                cCtrl.Parent.Controls(vA).Requery
            Next vA
            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrl.Parent, clObjCommButtons.sFrmMode, sFilGrp)
            
        
        Case sActType Like "Sort"
            Dim cList As Control
            Dim sSortedSQL As String
            Dim sSQL As String
            
            sFilGrp = clObjCommButtons.sFilGrp
            
            Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
            Set cList = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)
            
            'Obtém o SQL da [ ListBox ] do [ sFilGrp ]
            If InStr(cList.RowSource, "SELECT") > 0 Then sSQL = cList.RowSource Else sSQL = CurrentDb.QueryDefs(cList.RowSource).sql
            
            'Remove qualquer [ ORDER BY ] no SQL da lista
            If InStr(sSQL, "ORDER BY") > 0 Then sSQL = Split(sSQL, "ORDER BY")(0)
            
            'Aplica a ordenação no SQL da lista
            sSortedSQL = sSQL & "ORDER BY " & clObjCommButtons.sColToSort & " " & clObjCommButtons.sOrderMode
            
            'Atualiza o SQL da lista
            cList.RowSource = sSortedSQL
            
            'Inverte o [ sOrderMode ]
            If clObjCommButtons.sOrderMode = "ASC" Then
                If dictFndFldrVars("gBsSysLocalIcoZtoA") <> "" <> "" Then cCtrl.Picture = gBsSysLocalIcoAtoZ
                clObjCommButtons.sOrderMode = "DESC"
            Else
                If dictFndFldrVars("gBsSysLocalIcoAtoZ") <> "" <> "" Then cCtrl.Picture = gBsSysLocalIcoZtoA
                clObjCommButtons.sOrderMode = "ASC"
            End If
            
        Case sActType Like "UnHideAcc"
            Call Scr_ShowAccess
        
        Case Else
        
    End Select

End Sub


Private Sub mLabel_DblClick(Cancel As Integer)
    Dim fForm As Form
    Dim cCtrl As Control
    
    Set cCtrl = mLabel
    Set fForm = cCtrl.Parent
    
'Stop

    
    Call FormStatusBar04_OpnLogForm(fForm)
    
    
End Sub


Private Sub mListBox_BeforeUpdate(Cancel As Integer)

    Dim vA
    Dim iControlType As Integer
    Dim cCtrl As Control
    Dim sCtrL As String
    Dim fForm As Form

    Set cCtrl = mListBox
    Set fForm = cCtrl.Parent
    
    sCtrL = cCtrl.Name
    
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ] LostFocus, controle [ " & sCtrL & "]   "
'Stop

    iControlType = cCtrl.ControlType
    
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrl, True)

End Sub


Private Sub mTextBox_Change()

    Dim vA, vB
    Dim fForm As Form
    Dim cCtrl As Control

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Txtbox change"
'Stop
    Set cCtrl = mTextBox
    Set fForm = cCtrl.Parent
     
    ''-- Máscara de texto do controle ----
    ''Deve ser usado no evento Change do controle no próprio formulário
    '' pra que seja possível personalizar a máscara pra cada controle
    'Call MskdTxtbox02_TextMask(ActiveControl, ActiveControl.Text, "P_###,###", 6)
    
    'Aciona a atualização do [ TargetCtrl ] associado ao controle
    Call pb_TargtCtrlUpdate00_TimerDelay(fForm, cCtrl)
    

End Sub

Private Sub mTextBox_GotFocus()

    Dim iControlType As Integer
    Dim vA
    Dim cCtrl As Control
    Dim sCtrL As String
    Dim bLostFocus As Boolean
    
    'Dim fForM As Form
    
    Set cCtrl = mTextBox
    sCtrL = cCtrl.Name
    bLostFocus = False

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ]"
'Stop
    
    iControlType = cCtrl.ControlType
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrl, bLostFocus)

End Sub


Private Sub mTextBox_LostFocus()

    Dim vA
    Dim iControlType As Integer
    Dim cCtrl As Control
    Dim sCtrL As String
    Dim bLostFocus As Boolean
    'Dim fForM As Form
    
    Set cCtrl = mTextBox
    sCtrL = cCtrl.Name
    bLostFocus = False

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ]"
'Stop
    
    iControlType = cCtrl.ControlType
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrl, True)

End Sub


Private Sub mToggleButton_Click()
    Dim cCtrl As Control
    
    Set cCtrl = mToggleButton
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams
        Set .cTglBtnDocking = cCtrl
        Set .fTrgtForm = cCtrl.Parent
        .sTrgtForm = cCtrl.Parent.Name
        .bFrmIsDocked = .cTglBtnDocking.Value
        
    End With
    
    Call FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams)
    
End Sub











