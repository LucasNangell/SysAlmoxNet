Attribute VB_Name = "Módulo 07b - pbSubsTrggFilterBldSQL"
Option Compare Database
Option Explicit

Public Sub BuildSQL_CheckBox(cCtrL As Control, sTargtCtrlSQLselect As String, bMskdCtrl As Boolean)
    
    Dim vA, vB, vC
    Dim sOrigListTxt As String, sSrchTxt As String
    Dim iSrchVal As Integer
    
    Dim sReCntCpt As String
    Dim sSrchReCnt As String
    Dim sReCntFullStR As String
    Dim bNumImpar As Boolean
    Dim sCtrlQryField As String
    Dim sWhere As String
    Dim iSrchWildCard As String
    
    Dim sOpenBrkt As String, sCloseBrkt As String
    'dim sSrchVal as string

    vA = cCtrL.Name
    vB = cCtrL
'Stop
    'Roda o código apenas se o controle tiver algum item selecionado
    If Not IsNull(cCtrL) Then
        
        'Identifica os valores a serem usados na pesquisa (Numérico)
        '-------------------------------------------------------------------
        iSrchVal = cCtrL.Value     'Valor do CheckBox (Nulo: não selecionado / -1 selecionado / 0 não selecionado)
        sSrchReCnt = IIf(iSrchVal = -1, "Sim", "Não")
        '-------------------------------------------------------------------
'Stop
        'Monta o WHERE e o o RecCnt do controle
        '------------------------------------------------
        '---------------------------------------------------------------------------
        
        'Recupera o Campo de pesquisa na tabela
        sCtrlQryField = clObjTriggCtrlParam.sQryField
        '---------------------------------------------------------------------------
        '------------------------------------------------
        
        'Define como a expressão deve ser fechada dependendo se for um campo calculado ou não
        If clObjTriggCtrlParam.bBolClctd Then
            sOpenBrkt = ""
            sCloseBrkt = ""
        
        Else
            sOpenBrkt = "["
            sCloseBrkt = "]"
        
        End If
'Stop
        
        'Valor a ser pesquisado
        sWhere = sOpenBrkt & sCtrlQryField & " = " & iSrchVal & sCloseBrkt
        sWhere = "(" & sWhere & ")"
        
        
        'Texto a ser exibido no RecCntCpt
        sReCntCpt = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        sReCntFullStR = sReCntCpt & "[ " & sSrchReCnt & " ]"
        
        'Armazena os valores no objeto de classe [ clObjTriggCtrlParam ]
        clObjTriggCtrlParam.sGetSQLwhere = sWhere
        clObjTriggCtrlParam.sGetRecCntCptTxt = sReCntFullStR
        '-------------------------------------------------------------------
        
    End If
'Stop
End Sub


Public Sub BuildSQL_ComboBox(cCtrL As Control, sTargtCtrlSQLselect As String, bMskdCtrl As Boolean)
                              
    Dim vA, vB, vC
    Dim sOrigListTxt As String, sSrchTxt As String
    Dim iSrchVal As Integer
    
    Dim sReCntCpt As String
    Dim sSrchReCnt As String
    Dim sReCntFullStR As String
    Dim iItemPos As Integer
    Dim sCtrlQryField As String
    Dim sWhere As String
    Dim lngTbeClmn As Long
    Dim sOpenBrkt As String, sCloseBrkt As String
    
    vA = cCtrL.Name
'Stop

    lngTbeClmn = clObjTriggCtrlParam.iListboxTxtClmn
    
    'Roda o código apenas se o controle tiver algum item selecionado
    If Not IsNull(cCtrL.Value) Then
        
        'Identifica os valores a serem usados na pesquisa (Numérico)
        '-------------------------------------------------------------------
        iItemPos = cCtrL.ListIndex '(posição de tabela do item selecionado)
        sOrigListTxt = cCtrL.Column(lngTbeClmn, iItemPos)
        
        iSrchVal = cCtrL.Value     '(ID do item selecionado)
        'sSrchTxt = DesprezaAcentos(sOrigListTxt)
        sSrchReCnt = sOrigListTxt
        '-------------------------------------------------------------------
'Stop
        
        'Monta o WHERE e o o RecCnt do controle
        '------------------------------------------------
        '---------------------------------------------------------------------------
        
        'Recupera o Campo de pesquisa na tabela
        sCtrlQryField = clObjTriggCtrlParam.sQryField
        '---------------------------------------------------------------------------
        '------------------------------------------------
        
        'Define como a expressão deve ser fechada dependendo se for um campo calculado ou não
        If clObjTriggCtrlParam.bBolClctd Then
            sOpenBrkt = "("
            sCloseBrkt = ")"
        
        Else
            sOpenBrkt = "["
            sCloseBrkt = "]"
        
        End If
        
        
        'Valor a ser pesquisado
        sWhere = sOpenBrkt & sCtrlQryField & sCloseBrkt & " = " & iSrchVal
        sWhere = "(" & sWhere & ")"
        
        'Texto a ser exibido no RecCntCpt
        sReCntCpt = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        sReCntFullStR = sReCntCpt & "[ " & sSrchReCnt & " ]"
        
        'Armazena os valores no objeto de classe [ clObjTriggCtrlParam ]
        clObjTriggCtrlParam.sGetSQLwhere = sWhere
        clObjTriggCtrlParam.sGetRecCntCptTxt = sReCntFullStR
        '-------------------------------------------------------------------
        
    End If

'Stop
End Sub


Public Sub BuildSQL_ListBox(cCtrL As Control, sTargtCtrlSQLselect As String, bMskdCtrl As Boolean)

    Dim vA, vB, vC, vD, vE
    Dim sOrigListTxt As String, sSrchTxt As String
    Dim iSrchVal As Integer
    Dim sCtrL As String
    
    Dim sReCntCpt As String
    Dim sReCntFullStR As String
    
    Dim sCtrlQryField As String
    Dim sWhere As String
    Dim lngSelectedItems As Long
    Dim vListItem As Variant
    Dim vKey As Variant
    Dim lngCounT As Long
    Dim lngTbeClmn As Long
    Dim sOpenBrkt As String, sCloseBrkt As String
    
    sCtrL = cCtrL.Name
    vB = cCtrL
    
    'vA = cCtrL.ItemData(cCtrL.ListIndex)
    
MsgBox "teste --------------------------------------------------------------------------" & vbCr & "BuildSQL Listbox [ " & sCtrL & " ]"
Stop
    
    'clObjTriggCtrlParam.dictGetListSrchVals.RemoveAll
    'clObjTriggCtrlParam.dictGetListItemTxts.RemoveAll
    
    'Recupera a partir do objeto [ clObjTriggCtrlParam ] a coluna da tabela de dados da Listbox
    ' que tem o Valor Textual dos itens
    lngTbeClmn = clObjTriggCtrlParam.iListboxTxtClmn
    
    vA = cCtrL.ListIndex
    If vA > -1 Then cCtrL.Selected(vA) = True
    cCtrL.Value = cCtrL.Value
    vB = cCtrL.Value
    
    
    'Roda o código apenas se houver pelo menos um item selecionado na Lista
    lngSelectedItems = cCtrL.ItemsSelected.Count
'Stop
    
    If lngSelectedItems > 0 Then
        
        'Identifica os valores selecionados no Listbox (Numérico)
        '-------------------------------------------------------------------
        For Each vListItem In cCtrL.ItemsSelected     '-1 indica que o item está selecionado,  0 indica que não está
            
'Stop
            iSrchVal = cCtrL.ItemData(vListItem)                   'ID do item selecionado
            sOrigListTxt = cCtrL.Column(lngTbeClmn, vListItem)     'Texto associado ao item
            
            lngCounT = clObjTriggCtrlParam.dictGetListItemTxts.Count
            lngCounT = lngCounT + 1
            clObjTriggCtrlParam.dictGetListSrchVals.Add lngCounT, iSrchVal
            clObjTriggCtrlParam.dictGetListItemTxts.Add lngCounT, sOrigListTxt
            
        Next vListItem
        '-------------------------------------------------------------------
    
'    'Se o [ listbox ] for Multiselect = Nenhum [ ItemsSelected ] retorna ZERO
'    ' então é preciso recuperar o item selecionado de outra forma
'    Else
'        If cCtrL.ListIndex > -1 Then
''Stop
'            iSrchVal = cCtrL.Value                   'ID do item selecionado
'            sOrigListTxt = cCtrL.Column(lngTbeClmn, cCtrL.ListIndex)      'Texto associado ao item
'
'
'            lngCounT = clObjTriggCtrlParam.dictGetListItemTxts.Count
'            lngCounT = lngCounT + 1
'            clObjTriggCtrlParam.dictGetListSrchVals.Add lngCounT, iSrchVal
'            clObjTriggCtrlParam.dictGetListItemTxts.Add lngCounT, sOrigListTxt
        
'        End If
    
'    End If
    
    
'    If lngCounT > 0 Then
'Stop
        'Monta o WHERE e o o RecCnt do controle
        '------------------------------------------------
        '---------------------------------------------------------------------------
        
        'Recupera o Campo de pesquisa na tabela
        sCtrlQryField = clObjTriggCtrlParam.sQryField
        '---------------------------------------------------------------------------
        '------------------------------------------------
'Stop

        'Define como a expressão deve ser fechada dependendo se for um campo calculado ou não
        If clObjTriggCtrlParam.bBolClctd Then
            sOpenBrkt = "("
            sCloseBrkt = ")"
        
        Else
            sOpenBrkt = "["
            sCloseBrkt = "]"
        
        End If
        
        
        'Verifica os valores a serem filtrados
        lngCounT = clObjTriggCtrlParam.dictGetListSrchVals.Count
        For Each vKey In clObjTriggCtrlParam.dictGetListSrchVals
            lngCounT = lngCounT - 1
            
            If Not IsEmpty(clObjTriggCtrlParam.dictGetListSrchVals(vKey)) Then
'Stop
                'Valor a ser pesquisado
                vA = clObjTriggCtrlParam.dictGetListSrchVals(vKey)
                vC = sOpenBrkt & sCtrlQryField & sCloseBrkt & " = " & vA
                sWhere = sWhere & IIf(sWhere <> "", "", "(") & vC & IIf(lngCounT > 0, " or ", ")")
                
                'Texto a ser exibido no RecCntCpt
                vB = clObjTriggCtrlParam.dictGetListItemTxts(vKey)
                sReCntFullStR = sReCntFullStR & "[ " & vB & IIf(lngCounT > 0, " ] ou ", " ]")
            
            End If
'Stop
        Next vKey
'Stop
        'Texto a ser exibido no RecCntCpt
        sReCntCpt = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        sReCntFullStR = sReCntCpt & sReCntFullStR
        
        'Armazena os valores no objeto de classe [ clObjTriggCtrlParam ]
        clObjTriggCtrlParam.sGetSQLwhere = sWhere
        clObjTriggCtrlParam.sGetRecCntCptTxt = sReCntFullStR
        '-------------------------------------------------------------------
    
    End If
    
End Sub


Public Sub BuildSQL_OptionGroup(cCtrL As Control, sTargtCtrlSQLselect As String, bMskdCtrl As Boolean)
    
    Dim vA, vB, vC
    Dim sOrigListTxt As String, sSrchTxt As String
    Dim iSrchVal As Integer
    
    Dim sReCntCpt As String
    Dim sSrchReCnt As String
    Dim sReCntFullStR As String
    Dim bNumImpar As Boolean
    Dim sCtrlQryField As String
    Dim sWhere As String
    Dim iSrchWildCard As String
    Dim sOpenBrkt As String, sCloseBrkt As String

    vA = cCtrL.Name
'Stop
    
    'Verifica a quantidade de controles dentro do Opt Group
    vA = cCtrL.Controls.Count
    
    'Pra definir onde começa a contagem dos controles, verifica se o Grupo tem o rótulo
    'Se for ímpar significa que há um Rótulo associado ao controle
    ' Isso deve ser levado em consideração pra recuperar o Rótulo associado ao item selecionado
    vB = vA / 2
    bNumImpar = Int(vB) <> vB
'Stop
    'Roda o código apenas se o controle tiver algum item selecionado
    If Not IsNull(cCtrL) Then
        
        'Identifica os valores a serem usados na pesquisa (Texto associado à opção selecionada no controle)
        '--------------------------------------------------------------------------------------------------
        sOrigListTxt = cCtrL.Controls(IIf(bNumImpar, cCtrL * 2, cCtrL.Value + (cCtrL.Value - 1))).Caption
        sSrchTxt = DesprezaAcentos(sOrigListTxt)
        
        sSrchReCnt = sOrigListTxt
'Stop
        'Monta o WHERE e o o RecCnt do controle
        '------------------------------------------------
        '---------------------------------------------------------------------------
        
        'Recupera o Campo de pesquisa na tabela
        sCtrlQryField = clObjTriggCtrlParam.sQryField
        '---------------------------------------------------------------------------
        '------------------------------------------------
        
        'Define como a expressão deve ser fechada dependendo se for um campo calculado ou não
        If clObjTriggCtrlParam.bBolClctd Then
            sOpenBrkt = "("
            sCloseBrkt = ")"
        
        Else
            sOpenBrkt = "["
            sCloseBrkt = "]"
        
        End If
        
        'Valor a ser pesquisado
        iSrchWildCard = clObjTriggCtrlParam.iSrchWildCard
        vB = IIf(iSrchWildCard = 1, """" & sSrchTxt & "*""", IIf(iSrchWildCard = 2, """*" & sSrchTxt & "*""", """" & sSrchTxt & """"))
        sWhere = sOpenBrkt & sCtrlQryField & sCloseBrkt & " Like " & vB
        sWhere = "(" & sWhere & ")"
        
        'Texto a ser exibido no RecCntCpt
        vA = IIf(iSrchWildCard = 1 Or iSrchWildCard = 2, "*", "")
        vB = IIf(iSrchWildCard = 2, "*", "")
        sReCntCpt = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        sReCntFullStR = sReCntCpt & "[ " & vA & sSrchReCnt & vB & " ]"
        
        
        'Armazena os valores no objeto de classe [ clObjTriggCtrlParam ]
        clObjTriggCtrlParam.sGetSQLwhere = sWhere
        clObjTriggCtrlParam.sGetRecCntCptTxt = sReCntFullStR
        '-------------------------------------------------------------------
    End If
    
'Stop
End Sub

Public Sub BuildSQL_TextBox(cCtrL As Control, sTargtCtrlSQLselect As String, bMskdCtrl As Boolean)
    
    Dim vA, vB, vC
    Dim sOrigListTxt As String, sSrchTxt As String
    Dim iSrchVal As Integer
    
    Dim sReCntCpt As String
    Dim sSrchReCnt As String
    Dim sReCntFullStR As String
    Dim sCtrlQryField As String
    Dim sWhere As String
    Dim iSrchWildCard As String
    Dim sOpenBrkt As String, sCloseBrkt As String
    
    vA = cCtrL.Name
    vB = cCtrL.Value
    'vC = cCtrl.Text

    Debug.Print sTargtCtrlSQLselect
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "BuildSQL Textbox [ " & vA & " ]"
'Stop
    'Roda o código apenas se houver algum valor no controle
    
    On Error Resume Next
    'If gBbEnableErrorHandler Then On Error Resume Next
    
    'Se houver erro significa que o controle ora analisado não tem o foco
    ' nesse caso é preciso obter a proriedade .Value ao invés da .Text
    sOrigListTxt = cCtrL.Text
    If (Err.Number = 2185) Then
        sOrigListTxt = cCtrL.Value

    End If
    On Error GoTo -1
'Stop
    'verifica se o controle é do tipo Masked, ou seja, um textbox numérico que em tela é apresentado
    ' usando uma formatação de número cujos caracteres não devem ser considerados na pesquisa
    ' ex.: "#,###,###"  os caracteres , não devem ser descartados na filtragem
    ' no manual de parametrização procurar  [ Mskd1 ]  para mais informações
    ' ----
    'Pra confirmar que o parâmetro [ bMskdCtrl ] é true intencionalmente, só faz a limpeza
    ' caso a filtragem tiver sido disparada pelo evento [ Change ] do próprio controle
    If bMskdCtrl Then
'Stop
        vA = MskdTxtbox01_ClearNr(sOrigListTxt)
'Stop
        sOrigListTxt = vA
'        'Call MskdTxtbox02_TextMask(cTriggCtl, Me!txtProcesso.Text, "#,###,###", 7)
'        If sTxt <> "" Then vA = Val(CInt(sTxt))
'        sTxt = vA
    
    End If
    
    If sOrigListTxt <> "" Then
        
        'Identifica o valor a ser usado na pesquisa (Texto)
        '--------------------------------------------------------------------------------------------------
        sSrchTxt = DesprezaAcentos(sOrigListTxt)
        sSrchReCnt = sOrigListTxt
        '------------------------------------------------
'Stop
        'Monta o WHERE e o o RecCnt do controle
        '------------------------------------------------
        '---------------------------------------------------------------------------
        
        'Recupera o Campo de pesquisa na tabela
        sCtrlQryField = clObjTriggCtrlParam.sQryField
        '---------------------------------------------------------------------------
        '------------------------------------------------
'Stop
        'Define como a expressão deve ser fechada dependendo se for um campo calculado ou não
        If clObjTriggCtrlParam.bBolClctd Then
            sOpenBrkt = ""
            sCloseBrkt = ""
        
        Else
            sOpenBrkt = "["
            sCloseBrkt = "]"
        
        End If
        
        'Valor a ser pesquisado
        iSrchWildCard = clObjTriggCtrlParam.iSrchWildCard
        vB = IIf(iSrchWildCard = 1, """" & sSrchTxt & "*""", IIf(iSrchWildCard = 2, """*" & sSrchTxt & "*""", """" & sSrchTxt & """"))
        
        sWhere = sOpenBrkt & sCtrlQryField & sCloseBrkt & " Like " & vB
        sWhere = "(" & sWhere & ")"
'Stop
        'Texto a ser exibido no RecCntCpt
        vA = IIf(iSrchWildCard = 1 Or iSrchWildCard = 2, "*", "")
        vB = IIf(iSrchWildCard = 2, "*", "")
        
        vC = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        
        sReCntCpt = clObjTriggCtrlParam.sQryFieldCptClean & ": "
        sReCntFullStR = sReCntCpt & "[ " & vA & sSrchReCnt & vB & " ]"
        
        
        'Armazena os valores no objeto de classe [ clObjTriggCtrlParam ]
        clObjTriggCtrlParam.sGetSQLwhere = sWhere
        clObjTriggCtrlParam.sGetRecCntCptTxt = sReCntFullStR
        
        '-------------------------------------------------------------------
    End If
'Stop

End Sub
