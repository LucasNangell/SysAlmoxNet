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
Public Property Set InitCtrl(cCtrL As Control)
    Dim vA, vB
    Dim sCtrL As String

    'vA = bOnDirty
    sCtrL = cCtrL.Name
    vB = cCtrL.ControlType

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
    Select Case cCtrL.ControlType
        Case acCheckBox
            'Set mCheckBox = Nothing
            Set mCheckBox = cCtrL
            cCtrL.AfterUpdate = "[Event Procedure]"
            GoTo CaseResume     'por algum motivo quando o controle é um CheckBox o código
                                ' é direcionado pro Case Else ao invés de executar o acCheckBox e sair do Case
                                ' por esse motivo é necessário forçar sua saída
        
        Case acComboBox
            Set mComboBox = Nothing
            Set mComboBox = cCtrL
'Stop
            cCtrL.OnChange = "[Event Procedure]"
            cCtrL.AfterUpdate = "[Event Procedure]"

        Case acCommandButton
            Set mCommandButton = Nothing
            Set mCommandButton = cCtrL
            
            cCtrL.OnClick = "[Event Procedure]"

        Case acLabel
            Set mLabel = Nothing
            Set mLabel = cCtrL
            
            cCtrL.OnDblClick = "[Event Procedure]"

        Case acListBox
            Set mListBox = Nothing
            Set mListBox = cCtrL
            cCtrL.AfterUpdate = "[Event Procedure]"
            cCtrL.OnLostFocus = "[Event Procedure]"
            cCtrL.BeforeUpdate = "[Event Procedure]"
        Case acOptionGroup
'Stop
            Set mOptGroup = Nothing
            Set mOptGroup = cCtrL
            cCtrL.AfterUpdate = "[Event Procedure]"
            'GoTo CaseResume

        Case acTextBox
            Set mTextBox = Nothing
            Set mTextBox = cCtrL
'Stop
            cCtrL.OnChange = "[Event Procedure]"
            cCtrL.AfterUpdate = "[Event Procedure]"
            cCtrL.OnGotFocus = "[Event Procedure]"
            'cCtrl.BeforeUpdate = "[Event Procedure]"
            'cCtrl.OnKeyPress = "[Event Procedure]"
            cCtrL.OnLostFocus = "[Event Procedure]"
            
        Case acToggleButton
            Set mToggleButton = Nothing
            Set mToggleButton = cCtrL
            
            cCtrL.OnClick = "[Event Procedure]"
            
        Case Else
'Stop
            Exit Property
    
    End Select
'Stop
CaseResume:
'Stop
End Property


Private Sub mComboBox_AfterUpdate()
    Dim fForM As Form
    Dim cCtrL As Control
    
    Set cCtrL = mComboBox
    Set fForM = cCtrL.Parent
    
    'Aciona a atualização do [ TargetCtrl ] associado ao controle
    Call pb_TargtCtrlUpdate00_TimerDelay(fForM, cCtrL)
    

End Sub


Private Sub mCommandButton_Click()
    Dim vA
    Dim sMode As String, sForM As String, sActType As String, sFilGrp As String
    Dim cCtrL As Control, cEdtCtrl As Control
    Dim sCtrL As String
    Dim fForM As Form
    Dim sLstbxTag As String
    Dim vSplittedTAG As Variant
    Dim vKeyCtrl As Variant, vKeyBox As Variant
    Dim sBox As String
    Dim sQry As String, sQry2 As String
    
    
'MsgBox "testando RstArea commandbutton event"
'Stop
    On Error Resume Next
    Set cCtrL = mCommandButton
    sCtrL = cCtrL.Name
    Set fForM = cCtrL.Parent
    
    If Err.Number = 5 Then Exit Sub
    On Error GoTo -1
    
    If gBbEnableErrorHandler Then On Error Resume Next

    
    If Err.Number = 5 Then Exit Sub
    On Error GoTo -1

    sForM = cCtrL.Parent.Name
    
    Set clObjCommButtons = dictFormCommButtons(sForM)(cCtrL.Name)

    'Nos botões tipo [ RstArea ] o sActType retornado é "RstArea0#"
    ' o Select Case foi modificado pra permitir usar a função like
    ' e verificar se o início da string tem o termo "RstArea"
    sActType = clObjCommButtons.sActType
    
    Select Case True
        
        '--- CANCEL EDIT ----------------------------------------------
        Case sActType Like "Cancel*"
'Stop
            If TempVars("bReply") = True Then
                sFilGrp = clObjCommButtons.sFilGrp
                If sFilGrp <> "" Then
                    Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
                    sCtrL = clObjTargtCtrlParam.sTargtCtrlName
                    
                    If ControlExists(sCtrL, fForM) Then
                        Set cCtrL = fForM.Controls(sCtrL)
                        If fForM.Controls(sCtrL).ControlType = acListBox Then Call PbSubDataFields_FillFromListbox(cCtrL)
                        On Error GoTo -1
    'Stop
                    End If
                End If
                
                On Error GoTo -1
                Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)
                If gBbEnableErrorHandler Then On Error GoTo -1: On Error Resume Next

                                
                'código para recuperar em que [ Box ] o botão está inserido e executar a sub [ pbSub00_CtrlsEnblDsble_GetParams ]
                sBox = ""
                For Each vKeyBox In dictCtrlsInBox(sForM)
                    If dictCtrlsInBox(sForM)(vKeyBox) = clObjCommButtons.sCtrlName Then sBox = vKeyBox
                Next vKeyBox
                If sBox <> "" Then
                    Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, , sBox)
                Else
                    Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)
                End If
                
                'Mantém botão de edição habilitado caso haja algum item selecionado na lista
                If cCtrL.ItemsSelected.Count > 0 Then
                    Call pbSub00_CtrlsEnblDsble_GetParams(fForM, "RecSelect", sFilGrp)
                End If
                
                'Mantém botão de edição habilitado caso haja algum item selecionado na lista
                If cCtrL.ItemsSelected.Count > 0 Then
                    
                    Call pbSub00_CtrlsEnblDsble_GetParams(fForM, "RecSelect", sFilGrp)
                End If
            End If
            TempVars.RemoveAll
        
        '--- CANCEL NEW ----------------------------------------------
            If sActType = "CancelNew" Then
                sFilGrp = clObjCommButtons.sFilGrp
                If sFilGrp <> "" Then
                    For Each vKeyCtrl In dictFormDataFlds00Ctrls(sForM)
                        Set clObjCtrlDataFieds = dictFormDataFlds00Ctrls(sForM)(vKeyCtrl)
                        If clObjCtrlDataFieds.sDataFieldGrp = sFilGrp Then
                            If clObjCtrlDataFieds.sRecQry = "" Then
                                fForM.Controls(sCtrL).Value = ""
                            Else
                                sQry = fForM.Controls(sCtrL).RowSource
                                If InStr(sQry, "WHERE") > 0 Then
                                    If InStr(sQry, "GROUP BY") > 0 Then
                                        sQry2 = "GROUP BY " & Split(sQry, "GROUP BY")(1)
                                    ElseIf InStr(sQry, "ORDER BY") > 0 Then
                                        sQry2 = "ORDER BY " & Split(sQry, "ORDER BY")(1)
                                    End If
                                    sQry = Split(sQry, "WHERE")(0)
                                    sQry = sQry & vbCrLf & sQry2
                                    sQry = Replace(sQry, ";", "")
                                End If
                                fForM.Controls(sCtrL).RowSource = sQry
                            End If
                        End If
                    Next vKeyCtrl
                End If
                'Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)
            End If
        '--- CLOSE FORM ----------------------------------------------
        Case sActType Like "CloseForm"
            DoCmd.Close
        
        '--- EDIT REC ----------------------------------------------
        Case sActType Like "EditRec"
            sFilGrp = clObjCommButtons.sFilGrp
            If sFilGrp <> "" Then
                Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
                sCtrL = clObjTargtCtrlParam.sTargtCtrlName
    
                If ControlExists(sCtrL, fForM) Then
                    Set cCtrL = fForM.Controls(sCtrL)
                    If cCtrL.ControlType = acListBox Then
                    
                        If cCtrL.ListIndex > -1 Then
                            
                            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)
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
                      
        '--- HIDE ACCESS ----------------------------------------------
        Case sActType Like "HideAcc"
            'Oculta a janela do Access
            Call Scr_HideAccess
            
            'Traz o formulário pra frente pra reexibi-lo
            Call Scr_FormAlwaysOnTop(cCtrL.Parent, False)

        '--- NEW REC ----------------------------------------------
        Case sActType Like "NewRec"
'Stop
            sFilGrp = clObjCommButtons.sFilGrp
            If sFilGrp <> "" Then
                For Each vKeyCtrl In dictFormDataFlds00Ctrls(sForM)
                    Set clObjCtrlDataFieds = dictFormDataFlds00Ctrls(sForM)(vKeyCtrl)
                    If clObjCtrlDataFieds.sDataFieldGrp = sFilGrp Then
                        sCtrL = vKeyCtrl
                        If ControlExists(sCtrL, fForM) Then
                            If clObjCtrlDataFieds.sRecQry = "" Then
                                fForM.Controls(sCtrL).Value = ""
                            Else
                                sQry = fForM.Controls(sCtrL).RowSource
                                If InStr(sQry, "WHERE") > 0 Then
                                    If InStr(sQry, "GROUP BY") > 0 Then
                                        sQry2 = "GROUP BY " & Split(sQry, "GROUP BY")(1)
                                    ElseIf InStr(sQry, "ORDER BY") > 0 Then
                                        sQry2 = "ORDER BY " & Split(sQry, "ORDER BY")(1)
                                    End If
                                    sQry = Split(sQry, "WHERE")(0)
                                    sQry = sQry & vbCrLf & sQry2
                                    sQry = Replace(sQry, ";", "")
                                End If
                                fForM.Controls(sCtrL).RowSource = sQry
                            End If
                        End If
                        
                    End If
                Next vKeyCtrl
                
                For Each vKeyCtrl In dictFormCommButtons(sForM)
                    Set clObjCommButtons = dictFormCommButtons(sForM)(vKeyCtrl)
                    If clObjCommButtons.sFilGrp = sFilGrp And InStr(clObjCommButtons.sActType, "Save") > 0 Then clObjCommButtons.sActType = "SaveNew"
                Next vKeyCtrl
            End If
            
            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)

        '--- OPEN FORM ----------------------------------------------
        Case sActType Like "OpenForm"
            
            sFilGrp = clObjCommButtons.sFilGrp
            Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)
            sCtrL = clObjTargtCtrlParam.sTargtCtrlName
            Set cCtrL = fForM.Controls(sCtrL)
            
            'Atribuição dos valores a [ clObjFormOpenParams ] usados nas rotinas de abertura de formulários
            Set clObjFormOpenParams = New cls_09cParamsToOpenForms
            
            With clObjFormOpenParams
                
                'Garante que a variável [ bTrgtFormIsDocked ] que indica se o [ TrgtForm ] deve abrir ancorado
                ' receba False caso não tenha sido pressionada
                If Not IsNull(cCtrL.Parent.tglDockingTrgtFrm.Value) Then
                    Set .cTglBtnDocking = cCtrL.Parent.tglDockingTrgtFrm
                    If Not .cTglBtnDocking Is Nothing Then .bFrmIsDocked = .cTglBtnDocking.Value
                ElseIf ControlExists("tglDockingTrgtFrm", cCtrL.Parent) Then Set .cTglBtnDocking = cCtrL.Parent.tglDockingTrgtFrm
                End If
                .sFilGrp = sFilGrp
                .sFormMode = clObjCommButtons.sFrmMode
                .sTrggForM = cCtrL.Parent.Name
                .sTrgtForm = clObjCommButtons.sForM
                .lngRecID = cCtrL.ListIndex
                .bShowRecID = True
                .bSetPosition = True
                .bCentralizeForm = True
                Set .cTrgtCtrl = dictFormFilterGrpTrgts(cCtrL.Parent.Name)(sFilGrp)
                .sQryFieldID = GetsQryFieldID(clObjTargtCtrlParam)
            End With

                        
            'Inicia as rotinas pra abertura de formulário
            Call FormLoad01_FormLoadingStart(clObjFormOpenParams)
            
            'Abre o form [ frm_00(1)bSysTimer ] para monitorar "FormMove" tanto do TrggForm quanto do TrgtForm
            DoCmd.OpenForm "frm_00(1)bSysTimer", acNormal, , , , acHidden

        '--- RESET AREA ----------------------------------------------
        Case sActType Like "RstArea*"   '* é usado como Wildcard
'Stop
            Call pb_TargtCtrlUpdate04_RESETarea(cCtrL.Parent.Name, cCtrL.Name)

            
        '--- SAVE ----------------------------------------------
        Case sActType Like "Save*"      '* é usado como Wildcard
                            'o botão pode ser tipo [ SaveNew ] ou [ SaveEdit ]. a avaliação é feita na rotina [ PbSubDataFields_Rec ]
            sFilGrp = clObjCommButtons.sFilGrp
            Call PbSubDataFields_Rec(cCtrL)
            For Each vA In dictFormFilterGrpTrgts(sForM)(sFilGrp)
                cCtrL.Parent.Controls(vA).Requery
            Next vA
            Call pbSub00_CtrlsEnblDsble_GetParams(cCtrL.Parent, clObjCommButtons.sFrmMode, sFilGrp)
            
        
        '--- SORT ----------------------------------------------
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
                If dictFndFldrVars("gBsSysLocalIcoZtoA") <> "" <> "" Then cCtrL.Picture = gBsSysLocalIcoAtoZ
                clObjCommButtons.sOrderMode = "DESC"
            Else
                If dictFndFldrVars("gBsSysLocalIcoAtoZ") <> "" <> "" Then cCtrL.Picture = gBsSysLocalIcoZtoA
                clObjCommButtons.sOrderMode = "ASC"
            End If
            
        '--- UNHIDE ACCESS ----------------------------------------------
        Case sActType Like "UnHideAcc"
            Call Scr_ShowAccess
        
        Case Else
        
    End Select

End Sub


Private Sub mLabel_DblClick(Cancel As Integer)
    Dim fForM As Form
    Dim cCtrL As Control
    
    Set cCtrL = mLabel
    Set fForM = cCtrL.Parent
    
'Stop

    
    Call FormStatusBar04_OpnLogForm(fForM)
    
    
End Sub


Private Sub mListBox_BeforeUpdate(Cancel As Integer)

    Dim vA
    Dim iControlType As Integer
    Dim cCtrL As Control
    Dim sCtrL As String
    Dim fForM As Form

    Set cCtrL = mListBox
    Set fForM = cCtrL.Parent
    
    sCtrL = cCtrL.Name
    
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ] LostFocus, controle [ " & sCtrL & "]   "
'Stop

    iControlType = cCtrL.ControlType
    
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrL, True)

End Sub


Private Sub mTextBox_Change()

    Dim vA, vB
    Dim fForM As Form
    Dim cCtrL As Control

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Txtbox change"
'Stop
    Set cCtrL = mTextBox
    Set fForM = cCtrL.Parent
     
    ''-- Máscara de texto do controle ----
    ''Deve ser usado no evento Change do controle no próprio formulário
    '' pra que seja possível personalizar a máscara pra cada controle
    'Call MskdTxtbox02_TextMask(ActiveControl, ActiveControl.Text, "P_###,###", 6)
    
    'Aciona a atualização do [ TargetCtrl ] associado ao controle
    Call pb_TargtCtrlUpdate00_TimerDelay(fForM, cCtrL)
    

End Sub

Private Sub mTextBox_GotFocus()

    Dim iControlType As Integer
    Dim vA
    Dim cCtrL As Control
    Dim sCtrL As String
    Dim bLostFocus As Boolean
    
    'Dim fForM As Form
    
    Set cCtrL = mTextBox
    sCtrL = cCtrL.Name
    bLostFocus = False

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ]"
'Stop
    
    iControlType = cCtrL.ControlType
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrL, bLostFocus)

End Sub


Private Sub mTextBox_LostFocus()

    Dim vA
    Dim iControlType As Integer
    Dim cCtrL As Control
    Dim sCtrL As String
    Dim bLostFocus As Boolean
    'Dim fForM As Form
    
    Set cCtrL = mTextBox
    sCtrL = cCtrL.Name
    bLostFocus = False

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ]"
'Stop
    
    iControlType = cCtrL.ControlType
    'Chama a rotina de alteração de cor do controle
    Call HighlightClrChange(iControlType, cCtrL, True)

End Sub


Private Sub mToggleButton_Click()
    Dim cCtrL As Control
    
    Set cCtrL = mToggleButton
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams
        Set .cTglBtnDocking = cCtrL
        Set .fTrgtForm = cCtrL.Parent
        .sTrgtForm = cCtrL.Parent.Name
        .bFrmIsDocked = .cTglBtnDocking.Value
        
    End With
    
    Call FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams)
    
End Sub











