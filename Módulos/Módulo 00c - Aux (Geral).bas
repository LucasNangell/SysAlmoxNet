Attribute VB_Name = "Módulo 00c - Aux (Geral)"
Option Compare Database
Option Explicit

Public Function GetsQryFieldID(clObjTargtCtrlParam As cls_01aTargtCtrlParams_Evnts) As String
    Dim rsTbE As Recordset
    'Recupera o [ sQryIDfield ] do [ cCtrl ]
    ' [ sQryIDfield ] é o nome do campo utilizado como ID em uma consulta
    Set rsTbE = CurrentDb.OpenRecordset(clObjTargtCtrlParam.sClsLstbxSQL_eMAIN, dbOpenDynaset, dbReadOnly)
    GetsQryFieldID = rsTbE.Fields(0).Name
    rsTbE.Close
    Set rsTbE = Nothing

End Function
Public Function DialogBoxReply(sTitulo As String, sTexto As String, sTrggForM As String) As Boolean
    Dim vOpenArgs As Variant
    
    vOpenArgs = sTitulo & "|" & sTexto
'Stop
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    clObjFormOpenParams.sTrggForM = sTrggForM
    
    DoCmd.OpenForm "frm_00(1)dSysDialogBox", , , , , acDialog, vOpenArgs
    
    DialogBoxReply = TempVars("SysDialogBox")
    
'Stop
End Function



Public Sub TextboxScrollWhenNeeded(cCtrL As Control)

    Dim vA
    Dim dHgthLine As Double
    Dim lngVisibleLines As Long
    Dim lngQtTextLines As Long

    'Scroll bar only when needed
    dHgthLine = cCtrL.FontSize * 28
    lngVisibleLines = cCtrL.Height / dHgthLine

    'Scroll bar only when needed
    '---------------------------------------
    lngQtTextLines = UBound(Split(cCtrL.Value, vbCrLf))
    
    vA = IIf(lngQtTextLines > lngVisibleLines, 2, 0)
    cCtrL.ScrollBars = vA
    '---------------------------------------


End Sub



Public Function CreateMsgBoxTit(sModulo As String) As String
    
    'Define a 1a linha de uma Msgbox qualquer, geralmente utilizada pra título,
    ' usando a quantidade necessária de "-" pra preenchimento total da linha
    CreateMsgBoxTit = "----- " & sModulo & " "
    
    Do While Len(CreateMsgBoxTit) < 79
        CreateMsgBoxTit = CreateMsgBoxTit & "-"
    Loop
    
    CreateMsgBoxTit = CreateMsgBoxTit & vbCr & vbCr

End Function


Public Function GetTagParams(sParam As String, vTagSectionParams As Variant, _
                                    Optional cCascadeCtrL As Control, _
                                    Optional bErrorIfEmpty As Boolean, Optional vDefaulIfEmpty As Variant, Optional lngParamMin As Long, Optional lngParamMax As Long, _
                                    Optional bErrorOnScreen As Boolean, Optional sScreenStR1 As String, Optional sScreenStR2 As String, _
                                    Optional bErrorOnLog As Boolean, Optional sErrTitle As String, Optional cErrEvalCtrL As Control, Optional sLoadLogWarn As String) As Variant
    
    Dim vA, vB, vC
    Dim iWhere As Integer
    Dim iInT As Integer
    Dim vParamValue As Variant
    Dim bBoL As Boolean
    Dim sCascadeCtrL As String
    Dim sEvalCtrL As String
    Dim sForM As String
    
    
    'Dim bIsNumeric As Boolean
    
    'Alertas de erro:
    ' Erro de parâmetro vazio deve ser tratado fora da [ função ]
    ' . se o parâmetro ora avaliado estiver VAZIO ou não existir atribui [ vDefaulIfEmpty ] e sai da função
    '   nesse caso a mensagem de erro em tela deve ser chamada na saída da função
    
    ' Erro de parâmetro fora dos limites é tratado dentro da [ função ] com os parâmetros passados
    ' . se o valor definido no parâmetro estiver fora dos limites:
    '    . se [ sErrTitle ] for diferente de VAZIO o erro é carregado no log
    '    . se [ sScreenStR1 ] for diferente de VAZIO o erro é exibido na tela
    
    ' Erro de [ controle ] não localizado
    
    If Not cErrEvalCtrL Is Nothing Then
        sEvalCtrL = cErrEvalCtrL.Name
        sForM = cErrEvalCtrL.Parent.Name
    
    End If
    
    
    '------------------------------------------------------------------------------
    'Retorna o valor do parâmetro ora avaliado e se houver algum erro retorna VAZIO
    '------------------------------------------------------------------------------
    
    '------------------------------------------------------------------------------
    'Chamada da função
    'sParam = "Grp"
    '    sScreenStR1 = "Formulário:  [ " & sForm & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
    '    sScreenStR2 = " O [ grupo de filtragem ] do TargetCtrl não foi informado" & vbCr & "  e ele não poderá ser pesquisado."
    
    '    sFilGrp = GetTagParams(sParam, vTagSectionParams, vEmpty, lngMin, lngMax, sScreenStR1, sScreenStR2)
    '------------------------------------------------------------------------------
    
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Get TAG params function: [ " & sParam & " ]"
'Stop
    
    'Percorre cada um dos parâmetros da seção [ vTagSectionParams ]
    ' pra localizar o [ sParam ] ora avaliado
    iInT = 0
    Do
        iWhere = InStr(1, vTagSectionParams(iInT), sParam)
        iInT = iInT + 1
        
    Loop While iWhere = 0 And iInT <= UBound(vTagSectionParams)
    iInT = iInT - 1

'Stop
    'O parâmetro não foi localizado na TAG do controle
    If iWhere < 1 Then
        GetTagParams = vDefaulIfEmpty
        
        If bErrorIfEmpty Then GoTo EmptyValueHandler
        Exit Function
        
    End If

'Stop
    'Identifica o valor atribuído ao parâmetro
    vB = Len(vTagSectionParams(iInT)) - Len(sParam)
    vParamValue = Mid(vTagSectionParams(iInT), iWhere + Len(sParam), vB)
    
    If vParamValue = "" Then
        GetTagParams = vDefaulIfEmpty
        
        If bErrorIfEmpty Then GoTo EmptyValueHandler
        Exit Function
    
    End If
'Stop
    
    
    'Se na chamada da função tiver sido passado um [ TrggCtrl ], signfica que
    ' a variável ora avaliada deve retornar um controle
    ' portanto avalia se o controle recuperado no parâmetro existe no [ Form ]
    If Not cCascadeCtrL Is Nothing Then
    
        'Chama função pra confirmar a existência do controle
        ' e exibe alerta caso o RecCntCtrl indicado não exista
        sCascadeCtrL = vParamValue
        bBoL = ControlExists(sCascadeCtrL, cCascadeCtrL.Parent)
'Stop
        'Se o controle indicado nos parâmetros não for localizado exibe alerta de erro
        If Not bBoL Then
            GetTagParams = vDefaulIfEmpty
            
            vA = " Controle indicado nos parâmetros não foi localizado"
            '" Erro [ " & Err.Number & ": " & Err.Description & " ] "
            If bErrorOnScreen Then Call msgboxErrorAlert(sScreenStR1, sScreenStR2, vbExclamation, vA)
            If bErrorOnLog Then
                If sLoadLogWarn = "" Then sLoadLogWarn = "O controle indicado no parâmetro [ " & sParam & " ] dos seguintes controles não foi localizado."
                Call FormStatusBar01_Bld(sForM, sErrTitle, sLoadLogWarn, sEvalCtrL)
            End If
            
        Else
            GetTagParams = sCascadeCtrL
            
        End If
    
    'Se não tiver sido passado um controle deve retornar o valor do parâmetro ora avaliado
    Else
'Stop
        
            'Se não for um NÚMERO
            If Not IsNumeric(vParamValue) Then
                GetTagParams = vParamValue
                'GetTagParams = vDefaulIfEmpty
            
            Else
            
                'Se o parâmetro retornado for um NÚMERO
                ' e se tiver sido informado um [ valor máximo ] ou [ mínimo ], verifica se o valor retornado atende os requisitos e caso negativo:
                '  . exibe alerta
                '  . atribui o valor padrão e
                '  . ao sair dessa função, carrega o parâmetro no dicionário
    'Stop
                If (lngParamMax > 0 And vParamValue > lngParamMax) Or (lngParamMin > 0 And vParamValue < lngParamMin) Then
                    GetTagParams = vDefaulIfEmpty
                    
                    'Se tiver sido passado algum texto de erro exibe o alerta em tela
                    If bErrorOnScreen Then
                        vB = "Erro de parametrização de Controle do formulário"
                        Call msgboxErrorAlert(sScreenStR1, sScreenStR2, vbExclamation, vB)
    Stop
                    End If
                    
                    If bErrorOnLog Then
                        'sLoadLogWarn = sScreenStR1 & sScreenStR2
                        If sLoadLogWarn = "" Then sLoadLogWarn = "O parâmetro [ " & sParam & " ] dos seguintes controles não foi configurado com uma opção válida." & vbCrLf & "Os controles poderão não se comportar como esperado."
                        Call FormStatusBar01_Bld(sForM, sErrTitle, sLoadLogWarn, sEvalCtrL)
                    
                    End If
            
                Else
                    GetTagParams = Format(Val(vParamValue), "00")
                    
                End If
                
            End If
            
    End If
'Stop
Exit Function
        
EmptyValueHandler:
        'Se o parâmetro não existir ou não tiver sido definido
        ' e se houver indicação do título a ser carregado no log de carga do sistema

        If bErrorOnScreen Then
            vB = " -- "
            Call msgboxErrorAlert(sScreenStR1, sScreenStR2, vbExclamation, vB)
        End If

        If bErrorOnLog Then

            'sLoadLogWarn = sScreenStR1 & sScreenStR2
            If sLoadLogWarn = "" Then sLoadLogWarn = "O parâmetro [ " & sParam & " ] dos seguintes controles não foi configurado com uma opção válida." & vbCrLf & "Os controles poderão não se comportar como esperado."
            Call FormStatusBar01_Bld(sForM, sErrTitle, sLoadLogWarn, sEvalCtrL)

        End If
        
'Stop
End Function


Sub CleanDicts()

    Dim vA, vB, vC
    Dim dDicT1 As Dictionary, dDicT2 As Dictionary, dDicT3 As Dictionary, dDicT4 As Dictionary, dDicT5 As Dictionary, dDicT6 As Dictionary, dDicT7 As Dictionary
    
'Stop
    
    TempVars.RemoveAll

    '------------------------
    'sForM Dicts
    '------------------------
    dictCtrlBehvrParams.RemoveAll
    dictCtrlEnblDsblParams.RemoveAll     'já estava no Clean Dict
    dictCtrlsEvents.RemoveAll            'já estava no Clean Dict
    dictFormFilterGrps.RemoveAll         'já estava no Clean Dict
    dictFormsParams.RemoveAll
    dictFrmResetAreas.RemoveAll          'já estava no Clean Dict
    dictRstArBTNsByNr.RemoveAll
    dictRstArBTNsByName.RemoveAll
    dictTrgg00GrpsInForm.RemoveAll       'já estava no Clean Dict
    dictTrgg01CtrlsInGrp.RemoveAll       'já estava no Clean Dict
    dictTrggCtrlsInForm.RemoveAll        'já estava no Clean Dict
    dictTrgtCtrlsFilterGrps.RemoveAll
    dictFormQrysCtrls.RemoveAll
    dictFormCommButtons.RemoveAll
    dictFormDataFlds01Grps.RemoveAll
    dictFormFilterGrpsCtrls.RemoveAll
    '------------------------
    'já estavam no CleanDicts
    ' confirmar se estão sendo usados
    '------------------------
    dictCtrlTypeShort.RemoveAll
    dictCtrlTypeStR.RemoveAll
    
    
    dictParamByLckdStatus.RemoveAll
    Set dictParamByLckdStatus = Nothing
    
    'clObjUserParams.dictUserPermissions.RemoveAll
    Set clObjUserParams = Nothing
    
    dictUserPermissions.RemoveAll
    
    
    dictFormsParams.RemoveAll
    Set clObjStatusBarWarn = Nothing
    
    If Not clObjFormsParams Is Nothing Then clObjFormsParams.dForm_StatusBarWarns.RemoveAll
    Set clObjFormsParams = Nothing
        
        
    
'    Set clObjTargtCtrlParam = Nothing
'    Set clObjTriggCtrlParam = Nothing
'    Set clObjFilGrpsByForm = Nothing
'    Set clObjCtrlsEnblDsblParams = Nothing
'    Set clObjUserParams = Nothing


    '------------------------------------
    'inlcuir no Módulo Info
    

    '------------------------------------
'Stop

'Dicts e Classes publicas localizados por J Lucas


'dictCtrlBehvrParams.RemoveAll
'dictCtrlEnblDsblParams.RemoveAll
'dictCtrlsEvents.RemoveAll
'dictCtrlTypeShort.RemoveAll
'dictCtrlTypeStR.RemoveAll
'dictFormFilterGrps.RemoveAll
'dictFormsParams.RemoveAll
'dictFrmResetAreas.RemoveAll
'dictParamByLckdStatus.RemoveAll
'dictParamByLckdStatus.RemoveAll
'dictResetAreaCtrls.RemoveAll
'dictTrgg00GrpsInForm.RemoveAll
'dictTrgg01CtrlsInGrp.RemoveAll
'dictTrggCtrlsInForm.RemoveAll
'dictUserPermissions.RemoveAll


'Set clObjCtrlBehvrParams = Nothing
'Set clObjCtrlsEnblDsblParams = Nothing
'Set clObjCtrlsEvents = Nothing
'Set clObjFormsParams = Nothing
'Set clObjLckdStatusParam = Nothing
'Set clObjRstAreaParams = Nothing
'Set clObjStatusBarWarn = Nothing
'Set clObjTargtCtrlParam = Nothing
'Set clObjTrggCtrlGrpsByForm = Nothing
'Set clObjTriggCtrlParam = Nothing
'Set clObjUserParams = Nothing


End Sub


Function FindCodeLineInSub(sModName As String, sSubName As String, sSrchText As String) As Boolean
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim codeContent As Variant
    Dim iInT As Integer, iInT2 As Integer
    
    
    If gBbEnableErrorHandler Then On Error GoTo ErrorHandler
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "oMod: [ " & sSubName & " ] "
'Stop

    ' Acessa o projeto VBA
    Set vbProj = Application.VBE.ActiveVBProject
    Set vbComp = vbProj.VBComponents(sModName)
    Set codeMod = vbComp.CodeModule
    
    ' Obtém todo o código do módulo [ sModName ]
    codeContent = codeMod.Lines(1, codeMod.CountOfLines)
    
    'Separa as linhas do módulo [ sModName ]
    codeContent = Split(codeContent, vbCrLf)
    
    'Busca em cada linha do [ sModName ] pelo texto informado [ sSubName ]
    For iInT = 0 To UBound(codeContent) - 1
        If InStr(codeContent(iInT), sSubName) > 0 Then
            
            For iInT2 = iInT To UBound(codeContent) - 1
                
                If InStr(codeContent(iInT2), "End Sub") > 0 Then Exit Function
                    
                If InStr(codeContent(iInT2), sSrchText) > 0 Then
                    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "oMod found"
'Stop
                    FindCodeLineInSub = True
                    Exit Function
                    
                End If
        
            Next iInT2
            
        End If
        
    Next iInT
    
    Exit Function
    
ErrorHandler:
    MsgBox "Erro na busca pelo texto no módulo"
    
End Function


Public Function MskdTxtbox01_ClearNr(sTxtboxTxt As String) As String

    Dim vA, vB
    Dim sTxt As String
    Dim iInT As Integer

    '------------------------------------------
    'limpa o controle pra deixar apenas números
    If sTxtboxTxt <> "" Then
        For iInT = 1 To Len(sTxtboxTxt)
            sTxt = Mid(sTxtboxTxt, iInT, 1)
            If IsNumeric(sTxt) Then MskdTxtbox01_ClearNr = MskdTxtbox01_ClearNr & sTxt
        
        Next iInT
        
    End If
    
End Function





Function DesprezaAcentos(sTextoDigitado As String) As String
    Dim sVogaisSujas As String
    Dim sVogaisLimpas As String
    Dim sTempText As String
    'Dim sTextoDigitado As String
    Dim sTextoPuro As String
    Dim I As Long
    Dim sFinD As String
    Dim sReplaceBy As String
    Dim sTxtLst As String
    Dim sTxtStart As String
    Dim iTxtA As Integer, iTxtE As Integer, iTxtI As Integer, iTxtO As Integer, iTxtU As Integer, iTxtC As Integer, iTxtN As Integer
    Dim bAcentoFound As Boolean
    Dim iCompare As Integer

    'Código para permitir
'Stop

    'Define o método de comparação a ser usado para a substituição das letras acentuadas
    iCompare = vbTextCompare  'as letras com acentuação serão substituídas por suas equivalentes independentemente se estiverem em Cx Alta ou Cx Baixa
    
    'iCompare = vbBinaryCompare  'as letras com acentuação serão substituídas por suas equivalentes Cx Alta ou Cx Baixa
'Stop
    sTextoPuro = sTextoDigitado
    
    '-- versão antiga da pesquisa que apenas retira a acentuação do texto digitado pelo usuário mas
    '   não encontra o texto caso os dados na tabela pesquisada esejam acentuados --
    'Liste nesta variável todos os caracteres digitados pelo usuário
    ' que deverão ser substituídos para a realização da pesquisa
    sVogaisSujas = "àáâãäèéêëìíîïòóôõöùúûüñ"

    'Liste nesta variável todos os caracteres digitados pelo usuário
    ' que deverão ser substituídos para a realização da pesquisa
    sVogaisLimpas = "aaaaaeeeeiiiiooooouuuun"
    
    'Loop que percorrerá todas as letras da variável 'sVogaisSujas',
    'subtituindo os caracteres do texto digitado pelo usuário pelo caractere correspondente em 'sVogaisLimpas'
    For I = 1 To Len(sVogaisSujas)
        sFinD = Mid(sVogaisSujas, I, 1)
        sReplaceBy = Mid(sVogaisLimpas, I, 1)
        sTextoPuro = Replace(sTextoPuro, sFinD, sReplaceBy, , , iCompare)
'Stop
    Next I
    
'Stop
    
    'Trecho que inclui na string da SQL de pesquisa os eventuais acentos
    ' verifica a quantidade de vogais encontradas no texto
    iTxtA = Len(sTextoPuro) - Len(Replace(sTextoPuro, "a", ""))
    iTxtE = Len(sTextoPuro) - Len(Replace(sTextoPuro, "e", ""))
    iTxtI = Len(sTextoPuro) - Len(Replace(sTextoPuro, "i", ""))
    iTxtO = Len(sTextoPuro) - Len(Replace(sTextoPuro, "o", ""))
    iTxtU = Len(sTextoPuro) - Len(Replace(sTextoPuro, "u", ""))
    
    sTempText = sTextoPuro
    I = iTxtA + iTxtE + iTxtI + iTxtO + iTxtU
    If I > 0 Then
        If iTxtA > 0 Then
            sVogaisSujas = Replace(sTempText, "a", "[aàáâãä]", 1, -1, iCompare)
            sTempText = sVogaisSujas
        
        End If
        
        If iTxtE > 0 Then
            sVogaisSujas = Replace(sTempText, "e", "[eèéêë]", 1, -1, iCompare)
            sTempText = sVogaisSujas
        
        End If
        
        If iTxtI > 0 Then
            sVogaisSujas = Replace(sTempText, "i", "[iìíîï]", , , iCompare)
            sTempText = sVogaisSujas
        
        End If
        
        
        If iTxtO > 0 Then
            sVogaisSujas = Replace(sTempText, "o", "[oòóôõö]", , , iCompare)
            sTempText = sVogaisSujas
        
        End If
        
        
        If iTxtU > 0 Then
            sVogaisSujas = Replace(sTempText, "u", "[uùúûü]", , , iCompare)
            sTempText = sVogaisSujas
            
        End If
    
    End If
    
    'Retorna a string, convertida sem acentuação se for o caso
    DesprezaAcentos = sTempText

            
'Stop
End Function

Function GetQryFldOLD(sForM As String, sTrgtCtrl As String, sQryField As String) As vCheckQryFld
    Dim vA, vB, vC
    Dim rsTbE As Recordset
    Dim sQuerY As String
    Dim sWhere As String
    Dim lngFoundRecs As Long
    Dim fField As Field
    
    'Abre a consulta que será usada pra filtragem e confirma se o campo de consulta
    ' informado nos parâmetros do [ TriggCtrl ] existe
    sQuerY = Forms(sForM).Controls(sTrgtCtrl).RowSource
    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    
    For Each fField In rsTbE.Fields
        If fField.Name = sQryField Then GetQryFld.bFoundQryFld = True
                                                                                                                                                                
    Next fField
'Stop
    
    GetQryFld.sQry = sQuerY
'Stop
End Function


Sub msgboxErrorAlert(ByVal sMsgboxLine1 As String, Optional ByVal sMsgboxLine2 As String, Optional ByVal iButtons As Integer, Optional ByVal sTittle As String)
'Show Msgbox sub

    Dim sMsG As String
    
    sMsG = sMsgboxLine1 & IIf(sMsgboxLine2 <> "", vbCr & " " & sMsgboxLine2, "")
    MsgBox sMsG, iButtons, sTittle

    'Na rotina disparadora:
    
    'Dim sMsgboxLine1 As String, sMsgboxLine2 As String
    'sStR1 = " " & vbCr & " "
    'sStR2 = vbCr & " " & vbCr & " "
    
    'sStR1 = "Formulário:  [ " & sForM & " ]" & vbCr & "Listbox: " & " [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
    'sStR2 = " O controle [ " & sRecCntCtrl & " ] indicado como contador de registros" & vbCr & "  do TargtCtrl não foi localizado no formulário." & vbCr & "  A contagem de registros desse Listbox não será exibida."
    'vC = " Erro [ " & Err.Number & " ] "
    
    'Call msgboxErrorAlert(sMsgboxLine1, sMsgboxLine2, vbExclamation, vC)
'Stop
    'sStR1 = "": sGbMsgboxLine2 = ""
    
    'Botões
    '- vbInformation
    '- vbExclamation
    '- vbCritical
    '- vbQuestion

End Sub


Function ControlExists(sCtrL As String, fForM As Form) As Boolean
    Dim sTest As String
    
    'Testa se o Controle indicado existe
    If gBbEnableErrorHandler Then On Error Resume Next
    sTest = fForM(sCtrL).Name
    
    
    'A exmpressão Err.Number = 0 será falsa quando a tentativa de acessar
    ' o Controle informado resultar em erro, ou seja,
    ' ControlExists será True apenas com Err.Number for Zero
'Stop
    ControlExists = (Err.Number = 0)
    On Error GoTo -1
    
End Function


Public Function sGetClcltdField(sSQL As String, sField As String) As String

    Dim vA, vB, vC, vD
    Dim sSQLsize As String
    Dim lng_AS_pos As Long
    Dim lngStartingPos As Long
    Dim iWhere As Integer
    Dim sSQL_FrstSct As String
    Dim iSQL_FrstSctSize As Integer
    
    
    sSQLsize = Len(sSQL)  'comprimento do SQL
    '-----------------
    ' lngSQLsize = 216
    '-----------------
'Stop
    sSQL_FrstSct = "SELECT DISTINCT "
    iWhere = InStr(sSQL, sSQL_FrstSct)
    
    If iWhere > 0 Then
        
        'Prepara a Strig SQL
        '-----------------------------------------------------------------------------------
        iSQL_FrstSctSize = Len(sSQL_FrstSct)          'comprimento de SELECT DISTINCT
        sSQL_FrstSct = Left(sSQL, iSQL_FrstSctSize)   'Parte inicial da SQL: SELECT DISTINCT
        
        sSQL = Mid(sSQL, iSQL_FrstSctSize + 1, sSQLsize - iSQL_FrstSctSize)
        'Parte final da SQL: a partir de SELECT DISTINCT
        
        sSQL = ", " & sSQL
        If gBbDebugOn Then Debug.Print sSQL
        '-----------------
        '-----------------------------------------------------------------------------------
'Stop
        sSQLsize = Len(sSQL)
        '-----------------
        ' lngSQLsize = 202
        '-----------------
        
        lng_AS_pos = InStr(sSQL, sField)   'posição do campo a partir do início da SQL
        '-----------------
        ' lng_AS_pos  = 145    ", [tbl_02(3)cProdUnMedida_1].ProdUnidadeMedida AS ProdUnCons, [tbl_02(1)aProdutoBase].UnPedidoIDfk, [tbl_02(3)cProdUnMedida_2].ProdUnidadeMedida_"  (está na posição 145)
        '-----------------
        
        If gBbEnableErrorHandler Then On Error Resume Next
        lngStartingPos = InStrRev(sSQL, ", [", lng_AS_pos)
        If (Err.Number = 5) Then
            sGetClcltdField = "SELECT NotFound"
            Exit Function
            
        End If
        '-----------------
        ' lngStartingPos  = 1    ", [tbl_02(3)cProdUnMedida_1].ProdUnidadeMedida A
        '-----------------
        
        sGetClcltdField = Mid(sSQL, lngStartingPos + Len(", "), lng_AS_pos - (lngStartingPos + Len(", ")))
        
        '-----------------
        If gBbDebugOn Then Debug.Print "a" & sGetClcltdField
        
    Else
        sGetClcltdField = "SELECT NotFound"
    
    End If
    
'Stop
End Function


Public Function HexToLongRGB(sHexVal As String) As Long
'Obtém o valor de cor Decimal, tipo Long aceito nas propriedades de Cor dos controles,
' a partir do valor HEX color, valor exibido na Folha de Propriedades dos controles
'Stop
    Dim lRed As Long, lGreen As Long, lBlue As Long
    Dim vA, vB
    
    If Left(sHexVal, 1) = "#" Then sHexVal = Replace(sHexVal, "#", "")
    
    If sHexVal <> "0" Then
        'vA = ("&H" & Left$(sHexVal, 2))
        lRed = CLng("&H" & Left$(sHexVal, 2))      'left 2 chars are the red
        lGreen = CLng("&H" & Mid$(sHexVal, 3, 2))  'middle 2 are the green
        lBlue = CLng("&H" & Right$(sHexVal, 2))    'right 2 are the blue
    
    Else
        lRed = 0
        lGreen = 0
        lBlue = 0
    
    End If

    HexToLongRGB = RGB(lRed, lGreen, lBlue)
    'HexToLongRGB = RGB(CLng("&H" & Left$(sHexVal, 2)), CLng("&H" & Mid$(sHexVal, 3, 2)), CLng("&H" & Right$(sHexVal, 2)))
'Stop
End Function


