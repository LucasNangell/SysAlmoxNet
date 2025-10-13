Attribute VB_Name = "Módulo 07a - pbSubsTrggFilter"
Option Compare Database
Option Explicit

'----------------------------------------------------------------------
' Timer pra disparo da filtragem
'----------------------------------------------------------------------
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long
'----------------------------------------------------------------------
'-------------------------------------------

'Global I&

 Dim bProcessando As Boolean
 Dim lngTimerID As LongPtr
 '
 '


Public Sub pb_TargtCtrlUpdate00_TimerDelay(fForM As Form, cCtrL As Control)
'Public Sub pb_TargtCtrlUpdate00_TimerDelay()
    
    Dim vA, vB
    Dim sCtrL As String
    Dim sForM As String
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "TimerDelay trigger"
'Stop
    
    sForM = fForM.Name
    sCtrL = cCtrL.Name
    
    'Confirma se o controle é um [ TriggCtrl ]
    If IsObject(dictTrggCtrlsInForm(sForM)(sCtrL)) Then
        
        On Error Resume Next
        
        'Carrega as variáveis que serão usadas na rotina de atualização
        Set gBcTrggCtrl = cCtrL
        Set gBfTrggCtrlForm = fForM
        
'parei aqui1: verificar se há necessidade de "On Error..."
        
    
        'vA = gBcTrggCtrl.Name
        'vB = gBfTrggCtrlForm.Name
        
        
        'Set cCtrL = Screen.ActiveControl
    'Stop
        If lngTimerID <> 0 Then
                KillTimer 0, lngTimerID
                lngTimerID = 0
        End If
    'Stop
        'Chama a função pra atualizar a filtragem com base no Timer
        ' O Timer só funciona no modo de Execução ou pressionando o F5 em qualquer ponto do código, antes do encerramento da Rotina
        ' No F8, Modo Depuração o Timer não chega a zero e por isso a função não é chamada
        On Error GoTo -1
        lngTimerID = SetTimer(0, 0, gBiTypingDelay, AddressOf pb_TargtCtrlUpdate01_Start)
        On Error Resume Next
        
    'Stop
        If gBcTrggCtrl.ControlType = acTextBox Then
            gBcTrggCtrl.SetFocus
            gBcTrggCtrl.SelStart = Len(gBcTrggCtrl.Text)
    
        End If
    
    End If

End Sub

Public Sub pb_TargtCtrlUpdate01_Start()
    
    Dim vA, vB
'Stop
    'Impede múltiplas execuções simultâneas
    If bProcessando Then Exit Sub
    bProcessando = True
    
    ' Cancela o timer
    If lngTimerID <> 0 Then
        KillTimer 0, lngTimerID
        lngTimerID = 0
    
    End If
    
    'Confirma que o formulário está realmente aberto
    If CurrentProject.AllForms(gBfTrggCtrlForm.Name).IsLoaded Then
        
        
'Stop
    'vA = gBcTrggCtrl.Name
    'vB = gBfTrggCtrlForm.Name

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Update"
'Stop
        Call pb_TargtCtrlUpdate03_UNIQUEupdate(gBfTrggCtrlForm, gBcTrggCtrl)
    
    End If
    
    bProcessando = False
    
End Sub


Public Sub pb_TargtCtrlUpdate02_SetSearchType(fForM As Form, Optional cCtrL As Control, Optional sTagParams As String) ', Optional iResetArea As Integer)

    '-----------------------------------------------------------------------------
    ' Identifica o tipo de atualização disparada:
    ' .Reset Unique (atualização disparada a partir da alteração em um controle
    '    se tiver sido fornecido [ cCtrL ] com o controle de pesquisa que disparou a atualização
    '
    ' .Reset Area (atualização disparada a partir de um [ ResetBtn ]
    '    se tiver sido fornecido [ sTagParams ] indicando a ResetArea a ser atualizada
    '  e chama os procedimentos pra atualização das Listboxes adequadas
    '-----------------------------------------------------------------------------

' Avalia qual tipo de atualização foi disparado:
' - UniqueTrggCtrl
'    onde a atualização do Listbox foi disparada a partir do evento Change de um dos controles de filtragem
'     nesse caso devem ser identificados os TargtCtrls associados ao TriggCtrl individual.
'     A partir daí devem ser disparadas as atualizações dos TargtCtrls identificados

' - ResetBtn
'    onde a atualização do Listbox foi disparada a partir do pressionamento de um botão Reset
'     nesse caso devem ser identificados os TriggCtrls associados ao ResetBtn e em seguida
'     todos os TargtCtrls associados a esses controles. A partir daí devem ser disparadas as atualizações dos
'     TargtCtrls identificados


' Roteiro
'  -verificar no Dict do ResetBtn qual sua ResetArea pra identificar os TrggCtrls associados
'  -verificar no Dict do TriggCtrl qual o Grupo de Filtragem associado
'  -identificar o valor de cada um dos TrggCtrls associados àquele Grupo de Filtragem
'  -fazer a atualização dos TargtCtrls associados à area de filtragem identificada

End Sub


Public Sub pb_TargtCtrlUpdate03_UNIQUEupdate(fForM As Form, cCtrL As Control) ' , Optional sTagParams As String) ', Optional iResetArea As Integer)
    Dim vA, vB, vC
    Dim iResetArea As Integer
    Dim sForM As String
    Dim sCtrL As String
    Dim sFilGrp As String
    
    sForM = fForM.Name
    
'Stop
    'Confirma se foi fornecido um Controle na chamada da função
    If Not cCtrL Is Nothing Then
        sCtrL = cCtrL.Name

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Call pb_TargtCtrlUpdate06_BuildWHERE"
'Stop
        
        'Se dict [ dictTrggCtrlsInForm(sForM) ] não tiver sido carregado significa que o Form foi aberto
        ' sem a inicialização completa do sistema então não chama a atualização tipo Unique
        If IsObject(dictTrggCtrlsInForm(sForM)) Then
            
            'Confirma se o controle que disparou a alteração é um [ TriggCtrl ]
            ' do contrário não dispara a atualização de nenhum [ TrgtCtrl ]
            If IsObject(dictTrggCtrlsInForm(sForM)(sCtrL)) Then
                Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(sCtrL)
                sFilGrp = clObjFilGrpsByForm.sFilGrp
            
                '--------------------------------------------------------------------------------------------------------
                ' Quando a atualização for disparada a partir da alteração de um Controle no Form
                ' .identifica o Grupo de Filtragem do [ Trigg Control ] disparador e
                '  armazena no Dict [ dictTrgg01CtrlsInGrp ] os valores atuais de cada um dos controles
                '  associados e esse Grupo de Filtragem
                '--------------------------------------------------------------------------------------------------------
                On Error GoTo -1
                Call pb_TargtCtrlUpdate06_BuildWHERE(fForM, sFilGrp)
'Stop
            End If
            
        End If
'Stop
    End If
    
End Sub


Public Sub pb_TargtCtrlUpdate04_RESETarea(sForM As String, sResetAreaBtn As String)
    Dim vA, vB
    Dim sRstArea As String
    Dim vKeyCtrl As Variant, vKeyFilGrp As Variant
    Dim fForM As Form
    Dim sFilGrp As String
    Dim sKeyFilGrp As String
    Dim cCtrL As Control
    
    Set fForM = Forms(sForM)
    vA = fForM.Name
'Stop

    'Confirma se o dict [ dictRstArBTNsByNr(sForM) ] existe o que indica que há botões associados a [ Areas de Reset ]
    If IsObject(dictRstArBTNsByNr(sForM)) Then
'Stop
    
        'Acessa o dict [ dictRstArBTNsByName(sForM) ] para identificar a [ área de reset ] a ser esvaziada
        If dictRstArBTNsByName(sForM).Exists(sResetAreaBtn) = True Then
'Stop
            sRstArea = dictRstArBTNsByName(sForM)(sResetAreaBtn)
            
            
            'Confirma se o dict [ dictFrmResetAreas(sForM) ] com as [ áreas de reset ] do form com controles existe
            If IsObject(dictFrmResetAreas(sForM)) Then
                
                'Acessa o [ dictFrmResetAreas(sForM) ] da [ Area de Reset ] identificada
                ' pra localizar os [ controles ] que serão esvaziados
                If dictFrmResetAreas(sForM).Exists(sRstArea) Then
                
                    'Acessa a classe [ clObjRstAreaParams ] pra identificar
                    ' os [ Controles ] e os [ Grupos de Filtragem ] associados à [ Area de Reset ]
                    Set clObjRstAreaParams = dictFrmResetAreas(sForM)(sRstArea)
                    
                        'Por meio do dict [ clObjRstAreaParams.dictRstArCtrls ] identifica os controles da [ Área de Reset ] a serem esvaziados
                        For Each vKeyCtrl In clObjRstAreaParams.dictRstArCtrls
'Stop
                            Set cCtrL = Forms(sForM).Controls(vKeyCtrl)
                            
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Esvazia o controle [ " & vKeyCtrl & " ]"
'Stop
                            
                            'Esvazia o controle identificado
                            cCtrL = IIf(cCtrL.DefaultValue = "", Null, Replace(cCtrL.DefaultValue, """", ""))
                            Call HighlightClrChange(Int(cCtrL.ControlType), cCtrL, True)
                            
                        Next vKeyCtrl
                        
                        'Identifica os grupos de filtragem a [ Área de Reset ] pra atualizar
                        For Each vKeyFilGrp In clObjRstAreaParams.dictRstArFilGrps
                            
                            'Chama a atualização do [ Grupo de Filtragem ]
                            '--------------------------------------------------------------------------------------------------------
                            ' Quando a atualização for disparada por um botão de [ Reset ]
                            ' .identifica cada um dos Grupos de Filtragem associados à [ Reset Area ] e chama a atualização
                            '--------------------------------------------------------------------------------------------------------
                            On Error GoTo -1
                            sKeyFilGrp = vKeyFilGrp

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Atualiza o grupo [ " & sKeyFilGrp & " ]"
'Stop
                                  
                            Call pb_TargtCtrlUpdate06_BuildWHERE(fForM, sKeyFilGrp)
                        
                        Next vKeyFilGrp
                    
                End If
            Else
                MsgBox "erro não previsto"
                Stop
            
            End If
            
        Else
            MsgBox "erro não previsto"
            Stop
        
        End If
        
    Else
        MsgBox "erro não previsto"
        Stop
    
    End If

End Sub

Public Sub pb_TargtCtrlUpdate05_CleanCtrls_v1()
    Dim vA, vB, vC
    
Stop
    
    vA = dictCtrlTypeShort(cTweakableCtrL.ControlType)
    
    '-------------------------------------------------------------------------
    'Traz todos os Controles da Área de Reset indicada para seus
    ' valores padrão, conforme o tipo de controle
    '-------------------------------------------------------------------------
    Select Case sCtrlType
        Case "btn", "chk", "opb", "txt", "lst", "cmb"
    
    End Select

''    '--------------------------------------------------------------------------------------------------------
''    ' Identifica o Grupo de Filtragem do [ Trigg Control ] disparador e armazena no Dict [ dictTrgg01CtrlsInGrp ]
''    '  os valores atuais de cada um dos controles associados e esse Grupo de Filtragem
''    '--------------------------------------------------------------------------------------------------------



End Sub

Public Sub pb_TargtCtrlUpdate05_CleanCtrls_v2(iFrmIndexID As Integer, Optional iCtlRstArea As Integer, Optional bChgColor As Boolean)
    Dim vA, vB, vC
    Dim vKey As Variant
    Dim cCtL As Control
    Dim vListItem As Variant
    Dim iCtlType As Integer
    Dim iRstAr As Integer
    Dim bBoL As Boolean
    Dim dDicT As Dictionary


    Set dDicT = dictFilTriggCtlsByRstAr(iFrmIndexID, iCtlRstArea)
    bBoL = IsObject(dDicT)
    
    'retorna os controles a seus valores padrão
    If bBoL Then
        For Each vKey In dDicT
If Not bgbSkipStops4b Then Stop     'Ctrl point
            
            Set RstAreaCtls = dDicT(vKey)
            Set cCtL = RstAreaCtls.cCtrL
        
            If iCtlRstArea <> 0 Then
'Stop
                iCtlType = cCtL.ControlType
        

If Not bgbSkipMBox2a Then MsgBox "Reset de   [ " & vKey & " ]" & vbCr & vbCr & "---------------------- Fase quatro (Reset controls) ----------------------", vbOKOnly      'Mbox Ctrl point
If Not bgbSkipMBox2a Then Stop

If Not bgbSkipMBox4a Then MsgBox "Reset de   [ " & vKey & " ]" & vbCr & vbCr & "---------------------- Fase quatro (Reset controls) ----------------------", vbOKOnly      'Mbox Ctrl point
If Not bgbSkipStops4a Then Stop     'Ctrl point
        
                Select Case iCtlType
                    Case acCheckBox, acOptionGroup
'Stop
                        vA = cCtL.DefaultValue
                        cCtL.Value = vA
    
                    Case acListBox
'Stop
                        'remove de [ dictListboxCtrls(iFrmIndexID)(sListbox) / ListboxNav.dListboxItems ] todos os itens
                        ' previamente selecionados pelo usuário no Listbox
                        Set ListboxNav = dictListboxCtrls(iFrmIndexID)(cCtL.Name)
                        ListboxNav.dListboxItems.RemoveAll
                        
                        'remove a seleção de quaisquer itens que estejam selecionados no Listbox
                        For Each vListItem In cCtL.ItemsSelected   '-1 indica que o item está selecionado, 0 indica que não está
                            cCtL.Selected(vListItem) = False
    
                        Next vListItem
                        Call MudarCor(iCtlType, cCtL, True)
                    
                    Case Else
'Stop
                        vB = cCtL.Name
                        vA = cCtL.DefaultValue
                        cCtL.Value = vA
                        Call MudarCor(iCtlType, cCtL, True)
Stop
If Not bgbSkipStops4a Then Stop     'Ctrl point
                End Select
If Not bgbSkipStops4b Then Stop     'Ctrl point
                
            End If
        
        Next vKey
    
    End If
If Not bgbSkipStops4b Then Stop     'Ctrl point
'Stop
End Sub


Public Sub pb_TargtCtrlUpdate06_BuildWHERE(fForM As Form, sFilGrp As String)

    Dim vA, vB, vC, vD
    Dim sForM As String
    Dim sCtrL As String
    Dim cCtrL As Control
    Dim vKeyControl As Variant
    Dim iCtrlType As Integer
    Dim sSubToCall As String
    Dim lngCounT As Long
    Dim lngNonEmptyCTRLS As Long
    Dim sNewTrgtGrp_WHERE As String, sNewTrgtGrp_RecCntCpt As String
    Dim sJoint_WHERE As String, sJoint_RecCntCpt As String
    Dim sClose_WHERE As String, sClose_RecCntCpt As String
    Dim lngFilteredRecs As Long
    Dim lngRcstAllRecs As Long
    Dim lngDictItemsCnt As Long
    Dim sTargtCtrlName As String, sRecCntCtrlName As String
    Dim sTargtCtrlSQLselect As String
    Dim bBoL As Boolean
    Dim bMskdCtrlEventFound As Boolean
    Dim bMskdCtrl As Boolean
    Dim bActivateMask As Boolean
    Dim sStR1 As String, sStR2 As String
    Dim sLoadLogWarn As String
    Dim sCtrlEvent As String
    Dim sModName As String, sSubName As String, sSearchTerm As String
    Dim cTrggCtrl As Control
    Dim vKeyTrggCtrl As Variant
    
    sForM = fForM.Name
    
'MsgBox "----- pb_TargtCtrlUpdate06_BuildWHERE ----------------------------------------" & vbCr & vbCr & "Inicia BuildWHERE." & vbCr & " " & vbCr & " "
If gBbDepurandoLv03a Then Stop
'Stop

    '--------------------------------------------------------------------------------------------------------
    ' Passa por cada um dos controles do Grupo de Filtragem  do [ sForM ] e armazena no Dict [ dictTrgg01CtrlsInGrp ]
    '  os valores atuais de cada um dos controles associados e esse Grupo de Filtragem
    '--------------------------------------------------------------------------------------------------------
    
    'Se não houver [ TrggCtrls ] no [ grupo de filtragem ] atualmente avaliado não é necessário recuperar
    ' SQL e outros elementos já que não haverá filtragens
    If Not IsObject(dictTrgg01CtrlsInGrp(sFilGrp)) Then
        Exit Sub
        
    End If
    
    'Set dDicT = dictTrgg01CtrlsInGrp(sFilGrp)
    
    'If dictTrgg01CtrlsInGrp(sForM).Exists(sFilGrp) = True Then
    
    For Each vKeyControl In dictTrgg01CtrlsInGrp(sFilGrp)
        Set cCtrL = Forms(sForM).Controls(vKeyControl)
        vA = cCtrL.Name
        
        Set clObjTriggCtrlParam = dictTrgg01CtrlsInGrp(sFilGrp)(vKeyControl)
        'vB = cCtrl.Value
'Stop
        iCtrlType = cCtrL.ControlType
        sCtrL = cCtrL.Name

'MsgBox "----- pb_TargtCtrlUpdate06_BuildWHERE ----------------------------------------" & vbCr & vbCr & "1- Captura dados do controle [ " & sCtrL & " ]" & vbCr & "     do Grupo de Filtragem [ " & sFilGrp & " ]"
'If gBbDepurandoLv03a Then Stop
'Stop


        'Usa o dict [ dictCtrlTypeStR ] pra transformar, o Tipo do Controle identificado numericamente
        ' em texto indicando o tipo efetivo do controle
        ' e chamar a rotina correta pra guardar o valor atual de cada controle do Grupo de Filtragem
        sSubToCall = "BuildSQL_" & dictCtrlTypeStR(iCtrlType)
'Stop
        'Esvazia os dados armazenados em varreduras anteriores
        With clObjTriggCtrlParam
            .dictGetListSrchVals.RemoveAll
            .dictGetListItemTxts.RemoveAll
            .sGetSQLwhere = ""
            .sGetRecCntCptTxt = ""
        
        End With
'MsgBox "remove all"
'Stop
        
        'Recupera o SQL Select do TargtCtrl que está sendo atualizado,
        ' informação necessária para fazer a pesquisa em campos calculados
        Set clObjTargtCtrlParam = dictFormFilterGrps(sForM)(sFilGrp)
        sTargtCtrlSQLselect = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT
        
        'Debug.Print sTargtCtrlSQLselect
'Stop
        
        
If gBbDepurandoLv03a Then MsgBox "----- pb_TargtCtrlUpdate06_BuildWHERE ----------------------------------------" & vbCr & vbCr & "pass bMsked to TextBox filter: [ " & sCtrL & " ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv03a Then Stop
        
'Stop
        '--------------------------------------------------------------------------------
        '-----------------------------------------
        'Garante que apenas se o controle disparador cumprir os requisitos pra ser tratado como [ Masked ]
        ' será iniciada a rotina de descarte da máscara
        

If gBbDepurandoLv03a Then MsgBox "----- pb_TargtCtrlUpdate06_BuildWHERE ----------------------------------------" & vbCr & vbCr & "Avalia se há erros em [ BuildWhere ]: [ " & sCtrL & " ]"
If gBbDepurandoLv03a Then Stop
        
        'Verifica se o Controle existe no dict [ dictCtrlBehvrParams(sForM) ]
        ' e verifica se há erros referentes à carga do dicionário que devam ser registradas no Log de carga
        If dictCtrlBehvrParams(sForM).Exists(sCtrL) = True Then
            Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sCtrL)
'Stop
            bMskdCtrl = clObjCtrlBehvrParams.bMskdCtrl
            bMskdCtrlEventFound = clObjCtrlBehvrParams.bMskdCtrlEventFound
            
            'vB = clObjCtrlBehvrParams.bTriggrddByCtrlEvent
            
            'retorna TRUE se ambas as variáveis forem TRUE ou forem FALSE
            'retorna FALSE se as variáveis forem diferentes
            bBoL = bMskdCtrl = bMskdCtrlEventFound
            
            'Garante que [ bActivateMask ] terá o mesmo valor das variáveis [ bMskdCtrl ] e [ bMskdCtrlEventFound ]
            ' fazendo com que o evento Masked seja executado se TRUE e não executado se FALSE
            If bBoL Then
                bActivateMask = bMskdCtrl
            
            'Atribui FALSE a [ bActivateMask ] pra evitar que o código [ Masked ] seja executado
            ' e carrega pro log de carga do sistema o erro ocorrido conforme o caso
            Else
                bActivateMask = False
                
If gBbDepurandoLv03a Then MsgBox "----- pb_TargtCtrlUpdate06_BuildWHERE ----------------------------------------" & vbCr & vbCr & "Mask error em [ " & sCtrL & " ]"
If gBbDepurandoLv03a Then Stop
                
                If bMskdCtrl Then
                    vA = "Há controles com [ bMskdCtrl ] TRUE mas SEM a respectiva chamada [" & Chr(160) & "MskdTxtbox02_TextMask" & Chr(160) & "] no seu evento [" & Chr(160) & "Change" & Chr(160) & "]." & vbCrLf
                    vB = "Esses controles NÃO irão se comportar como [" & Chr(160) & "bMskdCtrl" & Chr(160) & "]."
                    
                    sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                    Call FormStatusBar01_Bld(sForM, "MskdMissingEvent", sLoadLogWarn, sCtrL)
                
                Else
                    vA = "Há controles COM a chamada [" & Chr(160) & "MskdTxtbox02_TextMask" & Chr(160) & "] no seu" & vbCrLf
                    vB = "evento [" & Chr(160) & "Change" & Chr(160) & "]" & " mas configurados com [ bMskdCtrl ] FALSE." & vbCrLf & "Esses controles NÃO irão se comportar como [" & Chr(160) & "bMskdCtrl" & Chr(160) & "]"
                    sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                    Call FormStatusBar01_Bld(sForM, "MskdMissingParam", sLoadLogWarn, sCtrL)
                
                End If
    
            End If
            
        End If
        '-----------------------------------------
        '--------------------------------------------------------------------------------
        

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Checa bMskdCtrl [ " & sCtrL & " ]"
'Stop
        vA = cCtrL.Name
        Application.Run sSubToCall, cCtrL, sTargtCtrlSQLselect, bActivateMask

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Back from Bld SQL [ " & sCtrL & " ]"
'Stop
        
        '-----------------------------------------
        '--------------------------------------------------------------------------------
        
        'Se o controle ora avaliado tiver algum valor é incluído
        ' na contagem de itens a serem incluídos no WHERE pra montagem final
        vA = clObjTriggCtrlParam.sGetSQLwhere
        lngNonEmptyCTRLS = IIf(vA <> "", lngNonEmptyCTRLS + 1, lngNonEmptyCTRLS)
        
    Next vKeyControl
        
'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "2- Encerrada a captura de dados dos controles" & vbCr & " do Grupo de Filtragem [ " & sFilGrp & " ]"
'Stop
        
    '--------------------------------------------------------------------------------------------------------
    'Passa por todos os [ TrggCtrls ] do Grupo de Filtragem
    ' e monta o WHERE e o RecCntCpt a partir dos valores guardados
    ' no objeto [ clObjTriggCtrlParam ]  da Classe [ cls_02aTrggCtrlParams ]
    ' de cada controle do Dict [ dictTrgg01CtrlsInGrp ]
    '--------------------------------------------------------------------------------------------------------
    
    lngDictItemsCnt = dictTrgg01CtrlsInGrp(sFilGrp).Count
    For Each vKeyControl In dictTrgg01CtrlsInGrp(sFilGrp)
        
        lngDictItemsCnt = lngDictItemsCnt - 1
        Set clObjTriggCtrlParam = dictTrgg01CtrlsInGrp(sFilGrp)(vKeyControl)

        vA = clObjTriggCtrlParam.sGetSQLwhere
        vB = clObjTriggCtrlParam.sGetRecCntCptTxt
        If gBbDebugOn Then Debug.Print vA

'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "3- Incluindo dados do Controle [ " & vKeyControl & " ]" & vbCr & "     na string WHERE e no RecCnt"
'Stop

        'Contagem de controles com valores a serem pesquisados
        If vA <> "" Then lngCounT = lngCounT + 1
        
        'Define se haverá inclusão de texto conector entre o trecho Anterior e o Atual
        If sNewTrgtGrp_RecCntCpt <> "" And vA <> "" Then
            sJoint_RecCntCpt = " AND "
            sJoint_WHERE = " / "
            
        End If
'Stop
        'Adiciona o novo trecho e inclui o WHERE nos parâmetros do Controle ora avaliado
        sNewTrgtGrp_WHERE = sNewTrgtGrp_WHERE & sJoint_RecCntCpt & vA
        sNewTrgtGrp_RecCntCpt = sNewTrgtGrp_RecCntCpt & sJoint_WHERE & vB
        sJoint_RecCntCpt = "": sJoint_WHERE = ""
        
        If gBbDebugOn Then Debug.Print sNewTrgtGrp_WHERE
'Stop
        'Se for o último WHERE a ser incluído,
        ' fecha a string de contagem de Registros
        '--------------------------------------------
        '----------------------------------------------------------------------------
'Stop
        'Só fecha a string WHERE se a quantidade de Controles verificados for igual
        ' à quantidade de controles que tem valores e
        ' se for o último Controle do Dict a ser verificado
        If lngCounT = lngNonEmptyCTRLS And lngDictItemsCnt = 0 Then
            
            'Fecha o WHERE do Grupo de Filtragem ora avaliado
            sNewTrgtGrp_WHERE = IIf(sNewTrgtGrp_WHERE <> "", "WHERE " & sNewTrgtGrp_WHERE, "")
            If gBbDebugOn Then Debug.Print sNewTrgtGrp_WHERE
            
'clObjTargtCtrlParam
            
            'Identifica os TargtCtrls que devem ser atualizados
            ' É necessário fazer a varredura no Dict pois é possível que
            ' um Grupo de Filtragem tenha mais de um TargtCtrl associado
            
        'For Each vKeyFilterGrp In dictFormFilterGrps(sForM) 'dictFormFilterGrps
            Set clObjTargtCtrlParam = dictFormFilterGrps(sForM)(sFilGrp)
            sTargtCtrlName = clObjTargtCtrlParam.sTargtCtrlName
            sRecCntCtrlName = clObjTargtCtrlParam.sRecCntCtrlName

            '-----------------------------------------------
            'Atualiza o TargtCtrl e o RecCnt
            '-----------------------------------------------
            If sTargtCtrlName <> "" Then  'Só atualiza se houver indicação do sTargtCtrlName que deve ser atualizado
                vA = sNewTrgtGrp_WHERE
'Stop
                Set cCtrL = Forms(sForM).Controls(sTargtCtrlName)
                'vA = cCtrl.Name
'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "4- atualiza o TargtCtrl e o RecCntCpt"
'Stop
                'Recupera o SQL do TargtCtrl
                vA = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT
                vB = clObjTargtCtrlParam.sClsLstbxSQL_bFROM
                vC = clObjTargtCtrlParam.sClsLstbxSQL_dOrderBy
                
                vD = vA & vbCr & vB & vbCr & sNewTrgtGrp_WHERE & vbCr & vC
                If gBbDebugOn Then Debug.Print vD
                
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Back from Bld SQL [ " & sCtrL & " ]"
'Stop
                
                vA = cCtrL.Name
                cCtrL.RowSource = vD
                
                'Se for uma Combobox e se o [ Trigger ] não estiver vazio exibe o primeiro item após o Trgt ser filtrado
                If cCtrL.ControlType = acComboBox Then

                    For Each vKeyTrggCtrl In dictTrgg01CtrlsInGrp(sFilGrp)
                        Set cTrggCtrl = Forms(sForM).Controls(vKeyTrggCtrl)
                        vA = cTrggCtrl.Name
                        
                        'Só exibe o primeiro item da lista da Combo se o [ Trigger ] que fez a filtragem não estiver vazio
                        On Error Resume Next
                        'Se houver erro significa que o controle ora analisado não tem o foco
                        ' nesse caso é preciso obter a proriedade .Value ao invés da .Text
                        vA = cTrggCtrl.Text
                        If (Err.Number = 2185) Then vA = cTrggCtrl.Value
                        On Error GoTo -1
                        
                        vA = IIf(vA = "", Null, vA)
                        
                        If Not IsNull(vA) Then cCtrL = cCtrL.ItemData(0) Else cCtrL.Value = Null
                        
                    Next vKeyTrggCtrl
                
                End If
                
                'Recupera a quantidade de registros exibidos
                ' apenas se tiver sido indicado um controle pra exibir
                bBoL = ControlExists(sRecCntCtrlName, fForM)
'Stop
                If sRecCntCtrlName <> "" And bBoL Then
                    
'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "5- retorna o total de registros da consulta do TargtCtrl"
'Stop
                    
'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "6- atualiza RecCnt"
'Stop
                    'Atualiza o RecCnt
                    
                    lngFilteredRecs = cCtrL.ListCount
                    Set cCtrL = Forms(sForM).Controls(sRecCntCtrlName)
                    vA = cCtrL.Name
                    
                    vA = IIf(lngFilteredRecs = 0, "Zero", Format(lngFilteredRecs, "#,###"))
                    vB = IIf(lngFilteredRecs = 1, ".", "s.")
                    
                    If sNewTrgtGrp_RecCntCpt = "" Then sNewTrgtGrp_RecCntCpt = "[ Todos os registros ]"
                    
                    'Se estiver vazio significa que não houve filtragem
                    sNewTrgtGrp_RecCntCpt = sNewTrgtGrp_RecCntCpt & " -> " & vA & " Reg" & vB
                    
                    cCtrL.Caption = sNewTrgtGrp_RecCntCpt
                
                End If
            
            End If
            
            '-----------------------------------------------
            '------------------------------

'        Next vKeyFilterGrp
        'clObjTargtCtrlParam

        End If
        '----------------------------------------------------------------------------
        '--------------------------------------------
        
    Next vKeyControl
    '----------------------------------------------------------------------------
    '--------------------------------------------
    
End Sub

