Attribute VB_Name = "Módulo 00a - Info"
Option Compare Database
Option Explicit


'-----------------------------------
'pendências
'-----------------------------------
' Ok -ResetArea

' Ok -Migração de eventos de controle pra Classe

' Ok -Dados (Exibição)

' Ok - checagem de arquivos de sistema no disco local
'        diferenciar erros de caminho e de ausência de arquivo


' Ok - abrir form Estoque com Edição desabilitada. Habilitar assim que um registro for selecionado
' Ok - ajustar EnableDisable pra habilitar [ btnEdtProd ] quando [ lstProdutos ] tiver um registro selecionado

' retirar edição de produtos da tela de Estoque. Haverá uma tela exclusiva pra isso.



' -Travar registro na rede --> OpenRecordset(strSQL, dbOpenDynaset, dbPessimistic)

' -Gravar dados (Edição, Inclusão)
' -Gerenciamento de Fotos

' -Serviços de manutenção do sistema



'-----------------------------------
'Melhorias
'-----------------------------------
' -Documentação
'  . Fluxograma
'  . Roteiro
'  . Parâmetros configuráveis
'    . TAGs de controles
'    . Constantes globais

' -Multiselect colocar em uso
' -CascUpDt  colocar em uso
' -Sub forms colocar em uso
' -Log de mudança de dados



' dictCtrlType.Add "104", "CommButton"
' dictCtrlType.Add "106", "CheckBox"
' dictCtrlType.Add "107", "OptionGroup"
' dictCtrlType.Add "109", "TextBox"
' dictCtrlType.Add "110", "ListBox"
' dictCtrlType.Add "111", "ComboBox"

'---------------------------------------------------------------------------------------------
' Pontos de parametrização do sistema
'---------------------------------------------------------------------------------------------
' . Constantes no módulo de Variáveis

' . Tabelas
'   . Niveis de permissão de acesso de usuários
'     [ tbl_00(2)bSysUserLoginLevels ]
'     . cadastro de perfis de usuário (ao cadastrar novos níveis de permissão lembrar que quanto menor o...)
       
'   . Permissões requeridas pra acesso a funcionalidades
'     [ tbl_00(3)aSysRqrdPrmss ]

'   . Parâmetros EnblDsbl para controles dos forms
'     [ qry_01(03)bSysEnblDisblParams ]

' . TAGs de controles
'   . Trigg... Trgt... Behv... ResetArea...
'   . SrtLbl...
'   . SrtBtn
'   . Reset buttons
'   . Multiselect...
'   . Data Field controls: DataLog... ChangesLog
'---------------------------------------------------------------------------------------------
    
    
    
'---------------------------------------------------------------------------------------------
' Sequência de carga de dicionários
'---------------------------------------------------------------------------------------------
    
'- Case acListBox, acComboBox
'  . pbSub21_TargtCtrlsDictBuild

'- Case acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox
'  . pbSub10_EventsDictBuild
'  . pbSub51_RstAreaDictBuild
'  . pbSub31_TriggCtrlDictPreBuild



'  - If iCtrlType = acListBox Or iCtrlType = acComboBox
'    . pbSub41_CtrlsBehvrDictBuild
    
'- Case acCheckBox, acOptionButton, acTextBox, acListBox, acComboBox, acCommandButton
'  . pbSub61_cCtrlsEnblDsblDictBuild
    
    
'---------------------------------------------------------------------------------------------
' Carga do sistema
'---------------------------------------------------------------------------------------------
' 1- Garante que os dicionários e objetos de classe estão vazios
'    Call CleanDicts

'---------------------------------------------------------------------------------------------

' '   ---------------------------------------------------------------------------------------------
' '2- Carrega o dicionário de CtrlTypes, que será usado em trechos do sistema
'     dictCtrlType.Add "104", "CommButton"
' '
' '   ---------------------------------------------------------------------------------------------
' '3- Identifica o usuário logado e recupera suas permissões de acesso
' '    Set clObjUserParams = New cls_08aLoggedUserParams
'      clObjUserParams.lngUserID = rsTbE.Fields("UserID")
'      ...
'      clObjUserParams.dictUserPermissions.Add CoDtxt, LoginLevelDescriç


' '   ---------------------------------------------------------------------------------------------
' '4- Percorre cada um dos Forms, abre em modo oculto e carrega os dicionários do sistema

' '   ---------------------------------------------------------------------------------------------
' '5- pbSub20_TargtCtrlsDictStartUp
        ' pra cada acListBox, acComboBox do [ Form ] avalia se o controle tem a 2a seção de parâmetros
        ' chama [ pbSub21_TargtCtrlsDictBuild ] pra carregar os parâmetros e
        ' monta o dicionário [ dictFormFilterGrpsTrgts(sForM) ] com os grupos de filtragem do [ Form ]
' '        . inclui no dict o grupo de filtragem associado ao controle ora avaliado
' '        . recupera o SQL do TrgtCtrl associado ao grupo
' '        . checa a inexistência de TrgtCtrls no form e inclui no dict de Log de Erros

' '   ---------------------------------------------------------------------------------------------
' '6- pbSub30_TriggCtrlDictStartUp
        ' se o Form tiver TrgtCtrls passa por todos os acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox do Form
        ' e avalia se o controle tem a TAG necessária pra inclusão nos dicts" & vbCr & "[ dictTrgg00GrpsInForm ] e [ dictTrgg01CtrlsInGrp ]
        ' Avalia se o controle é um dos tipos da categoria BehvrCtrls (acTextBox, acListBox, acComboBox)
        ' Avalia se o controle tem a TAG necessária pra inclusão no dict [ dictCtrlBehvrParams(sForM) ]
        ' e caso postivio inclui o controle no dict
        '
        '-
        '-
'pbSub10_EventsDictBuild
'pbSub31_TriggCtrlDictPreBuild
'pbSub41_CtrlsBehvrDictBuild


' '7- Avalia erros de carga do sistema: controle não configurado como [ Trigger] mas com [ bOnDirty ] TRUE


' '        .
' '        .
' '        .
' '        .
' '   -
' '   -
' '   -
' '   -
' '   -
' '   -





' '  ----------------------------------------------------------------
' '  Checagem de Erros
' '  ---------------------------------------------------------------------------------------------
' '  ---------------------------------------------------------------------------------------------
' '  1. Durante a carga do sistema se encontrado algum erro é chamada a rotina [ FormStatusBar01_Bld ]
' '   pra montagem do [ dictFormsParams ] com informações sobre falhas na carga de parâmetros de controles do [ Form ]
' '
' '  2. Ao final da carga do sistema se o parâmetro [ bForm_ShowWarns ] do dict [ dictFormsParams ] for TRUE
' '     chama a rotina [ FormStatusBar02_SetWarn(fForM) ] pra alterar a Status Bar do Form recém carregado
' '     e exibir o alerta de erros na carga do sistema
' '
' '  3. A rotina [ FormLoad06b_BackFromFormLoad(sForM) ] é chamada quando for necessário avaliar se [ bForm_ShowWarns ] é TRUE

' '  4. Após a abertura do Form se a [ StatusBar ] receber um duplo clique
' '     é chamada a rotina [ FormStatusBar04_OpnLogForm ] pra montar o texto a ser exibido
' '     no Form [ frm_00(1)cSysLoadLog ] indicando o Log de alertas de carga do sistema
' '
' '
' '
' '
' '

' '
' '  FormStatusBar01_Bld(sForM, "...", sLoadLogWarn)
' '  ... MissingBehvrParams
' '   "Há controles no Formulário sem [ BehvrParams ]. Esses controles não irão se comportar como esperado."
' '
' '  ... NoTrgtCtrls
' '   "Não foram encontrados TargetCtrls no formulário. Não será possível fazer pesquisas."
' '
' '  ... DirtyTrue_NoTrgg
' '   "Há controles não carregados como [ Trigger ] mas configurados com [ OnDirty ]"
                
                
                
' '  ----------------------------------------------------------------
' '  Roteiro pra disparo de pesquisas em [ TriggCtrls ]
' '  ---------------------------------------------------------------------------------------------
' '  1. Disparado o evento [ Change ] do próprio controle
' '
' '  2. Chama a rotina [ MskdTxtbox02_TextMask ] pra tratar o campo
' '     de forma a aplicar a máscara definida pro controle
' '

' '     substituído por [ bMskdCtrlEventFound ] apenas dentro da rogina BuildWhere
' '     Atribui TRUE à variável [ bTriggrddByCtrlEvent ] pra garantir que
' '     o código de pesquisa desconsidere a máscara do controle

' '  3. Segue pro evento Change da Classe [ cls_10aCtrls_Events ] de eventos de controle
' '  4. Chama a rotina [ pb_TargtCtrlUpdate00_TimerDelay ] pra iniciar o timer de atualização do controle
' '  5. Chama a rotina [ pb_TargtCtrlUpdate01_Start ] pra iniciar a atualização do controle
' '  6. Chama a rotina [ pb_TargtCtrlUpdate03_UNIQUEupdate ] pra identificar:
' '      - o Grupo de Filtragem do controle disparador
' '      - o tipo de atualização disparada:
' '        . Reset Unique
' '        . Reset Area
' '  8. Chama a rotina [ pb_TargtCtrlUpdate06_BuildWHERE ] pra iniciar a montagem da cláusula WHERE pra atualização do [ TrgtCtrl ]
' '       Se o controle não tiver sido disparado pelo próprio evento [ Change ] trata o controle
' '       como se fosse [ Masked ] FALSE
' '  9. Chama a rotina [ BuildSQL (_TextBox) ] apropriada a depender do tipo de controle disparador
' '      se a configuração [ Masked ] do controle for TRUE, significa além de ter TRUE no parâmetro
' '      a pesquisa foi disparada pelo evento [ Change ] do próprio controle, o que confirma que ele é um [ Masked ]
' '  ----------------------------------------------------------------
                
                

' '  ----------------------------------------------------------------
' '  Roteiro de montagem de dicionários
' '  ---------------------------------------------------------------------------------------------
' '
' '   ---------------------------------------------------------------------------------------------
' '   Percorre cada um dos Forms, abre em modo oculto e carrega os dicionários do sistema
' '    em seguida fecha o Form (os objetos carregados nas classes são perdidos ao fechar o Form
' '    então descartei o carregamento direto de objetos)
    
' '1- carrega os dados dos TargtCtrls (Listbox e Combobox)
' '  Tratamentos de erro (CHECAGENS)
' '   - Parâmetro RecCntCtrl vazio
' '      . Ao executar uma filtragem o sistema não tenta exibir a contagem de registros
' '        se o parâmetro for vazio
' '
' '   - RecCntCtrl associado ao TargtCtrl não existente no form
' '      . Exibe alerta, continua a carga do controle mas deixa o parâmetro RecCntCtrl vazio.
' '        Assim evita que o sistema tente encontrá-lo pra exibir a contagem de registros após filtragem.
' '
' '   - Mais de um TargtCtrl associado ao mesmo Grupo de Filtragem
' '      . Dá alerta e não carrega o Listbox como TargtCtrl
' '
' '. pbSub20_TargtCtrlsDictStartUp
' ' . pbSub21_TargtCtrlsDictBuild
' '    Cria uma nova variação do dicionário pro Formulário corrente
'       Set dictFormFilterGrpsTrgts(sForm) = New Dictionary
'       Set dDict = dictFormFilterGrpsTrgts(sForm)
' '
' '   Recupera o Grupo de Filtragem pra montar o Dicionário
' '   . pbSub06_GetCtrlTagFltrGrp
' '
' '   Cria um novo objeto [ clObjTargtCtrlParam ] da Classe [ cls_01aTargtCtrlParams_Evnts ]
' '    pra ser incluído no Dict [ dictFormFilterGrpsTrgts(sForm) ]
'      Set clObjTargtCtrlParam = New cls_01aTargtCtrlParams_Evnts
' '
' '   Adiciona um novo item no dicionário [ dictFormFilterGrpsTrgts ] e guarda nele o objeto [ clObjTargtCtrlParam ]
' '    com os respectivos parâmetros do targtCtrl definidos na classe [ cls_01aTargtCtrlParams_Evnts ]
'       dDict.Add sFilGrp, clObjTargtCtrlParam
' '
' '    Obs: não foi prevista a possibilidade de incluir num Grupo de Filtragem mais de um TargtCtrl
' '          pra isso funcionar, ao invés de incluir no [ dictFormFilterGrpsTrgts(sForm) ] um objeto de classe
' '          com os parâmetros do Grupo (TargtCtrl, Where, etc), teria que ser incluído um segundo dicionário.
' '          Nesse segundo dicionário seriam incluídos cada um dos TargtCtrls associados ao Grupo, e nesse dicionário
' '          aí sim seria incluído o [ clObjTargtCtrlParam ]
' '
' '     Atribui ao Listbox os parâmetros esperados pela Classe [ cls_01aTargtCtrlParams_Evnts ]
' '     -------------------------------------------------------------------------------------------
' '      dictFormFilterGrpsTrgts   - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictFormFilterGrpsTrgts(sForm)     - Grupos de Filtragem
' '          Item "01"  >  clObjTargtCtrlParam  - parâmetros do Grupo de Filtragem
' '                         .sTargtCtrlName
' '                         .sClsLstbxSQL_aSELECT
' '                         .sClsLstbxSQL_bFROM
' '                         .sClsLstbxSQL_cWHERE
' '                         .sClsLstbxSQL_dOrderBy
' '                         .sClsLstbxSQL_eMAIN
' '                         .sFilGrp
' '                         .sRecCntCtrlName
' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------
'      Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)("01")
'      va = clObjTargtCtrlParam.sTargtCtrlName
'
'      For Each vKey In dictFormFilterGrpsTrgts(sForM)
'          Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(vKey)
'
'      Next vKey
'     ---------------------------------------------------------------------------------------------
' '
' '2a- carrega os dados dos TriggCtrls
' '. pbSub30_TriggCtrlDictStartUp(sForm)
' ' . pbSub31_TriggCtrlDictPreBuild(sCtrlTAG, cTriggCtrl)

' '   Recupera o Grupo de Filtragem pra montar o Dicionário
' '   . pbSub06_GetCtrlTagFltrGrp

' '   Cria o dicionário com grupos de filtragem pra cada formulário
'     Set dictTrgg00GrpsInForm(sForm) = New Dictionary

' '   Cria o nested Dict com os controles associados a cada Grupo de Filtragem
'      Set dictTrgg01CtrlsInGrp(sFilGrp) = New Dictionary
'      Set dDict = dictTrgg01CtrlsInGrp(sFilGrp)
 
' '   Cria um novo objeto [ clObjTriggCtrlParam ] da Classe [ cls_02aTrggCtrlParams ]
' '    pra ser incluído no Dict [ dictTrgg01CtrlsInGrp(sFilGrp) ]
'       Set clObjTriggCtrlParam = New cls_02aTrggCtrlParams
'       dDict.Add sCtrL, clObjTriggCtrlParam
' '   ---------------------------------------------------------------------------------------------
' '      dictTrgg00GrpsInForm  - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictTrgg00GrpsInForm(sForM)     - Grupos de Filtragem
' '          Item "01"   > dictTrgg01CtrlsInGrp(sFilGrp) - Controles associados ao Grupo
' '                          Item "cmbSrcCateg"  >  clObjTriggCtrlParam  - parâmetros do Controle
' '                                                  .sQryField = sQryField
' '                                                  .sFilGrp = sFilGrp
' '                                                  .iQryFldRmvCharCpt = iQryFldRmvCharCpt
' '                                                  .iSrchWildCard = iSrchWildCard
' '                                                  .iClctdStrSze = iClctdStrSze
' '                                                  .iSrchOnChange = iSrchOnChange
' '                                                  .sCascUpDtTrgCtrl = sCascUpDtTrgCtrl
' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------
'      Set clObjTriggCtrlParam = dictTrgg00GrpsInForm(sForM)("01")("cmbSrcCateg")
'      va = clObjTriggCtrlParam.sQryField
'
'      For Each vKey In dictTrgg00GrpsInForm(sForM)
'          For Each vA In dictTrgg01CtrlsInGrp(vKey)
'
'              Set clObjTriggCtrlParam = dictTrgg01CtrlsInGrp(vKey)(vA)
'                  vB = clObjTriggCtrlParam.sCtrlName
'
'          Next vA
'
'      Next vKey
'     ---------------------------------------------------------------------------------------------

 
' '2b- Cria o dicionário com os Trigg Controls de cada formulário, indicando o Grupo de Filtragem associado a cada controle
' '     ainda no pbSub31_TriggCtrlDictPreBuild
 
'      Set dictTrggCtrls_FilGrp(sForM) = New Dictionary
'      Set dDict = dictTrggCtrls_FilGrp(sForM)
'      Set clObjFilGrpsByForm = New cls_02bTrggCtrlGrpsByForm

'      dDict.Add sCtrL, clObjFilGrpsByForm
'      Set clObjFilGrpsByForm.cCtrL = cTriggCtrl
'      clObjFilGrpsByForm.sFilGrp = sFilGrp
' '   ---------------------------------------------------------------------------------------------
' '      dictTrggCtrls_FilGrp  - Forms
' '        Item "frm_01(1)bProdEstoque"
'
' '        dictTrgg00GrpsInForm(sForM)     - Controles do Form
' '          Item "cmbSrcCateg"  > clObjFilGrpsByForm  - parâmetros do Controle
' '                                   .sCtrlName
' '                                   .sFilGrp
' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------
'      Set clObjFilGrpsByForm = dictTrggCtrls_FilGrp(sForM)("cmbSrcCateg")
'      vA = clObjFilGrpsByForm.sCtrlName
'
'      For Each vKey In dictTrggCtrls_FilGrp(sForM)
'              Set clObjFilGrpsByForm = dictTrggCtrls_FilGrp(sForM)(vKey)
'              vA = clObjFilGrpsByForm.sCtrlName
'
'      Next vKey
'     ---------------------------------------------------------------------------------------------

' '3- carrega os dados das ResetAreas
' '. pbSub50_RstAreaDictStartUp(fForm)
' ' . pbSub51_RstAreaDictBuild(sCtrlTag, cRstAreaCtrL)

' '   Recupera a ID da ResetArea pra montar o Dicionário

' '   Cria o dicionário com Reset Areas pra cada formulário
'     Set dictFrmResetAreas(sForM) = New Dictionary
'     Set dDictOuter = dictFrmResetAreas(sForM)

' '    Cria o nested Dict com os controles associados a cada ResetArea
'      Set dictResetAreaCtrls(sRstArea)= New Dictionary
'      Set dDictInner = dictResetAreaCtrls(sRstArea)
 
' '    Cria um novo objeto [ clObjRstAreaParams ] da Classe [ cls_05aResetAreasParams ]
' '     pra ser incluído no Dict [ dictResetAreaCtrls(sRstArea) ]
'      Set clObjRstAreaParams = New cls_05aResetAreasParams
'      dDictInner.Add sCtrl, clObjRstAreaParams
' '   ---------------------------------------------------------------------------------------------
' '      dictFrmResetAreas  - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictFrmResetAreas(sForM)     - Áreas de Reset
' '          Item "01"   > dictResetAreaCtrls(sRstArea) - Controles associados à Área de Reset
' '                          Item "cmbSrcCateg"  >  clObjRstAreaParams  - parâmetros do Controle
' '                                                  .sQryField = sQryField
' '                                                  .sRstAr = sRstAr
' '   ---------------------------------------------------------------------------------------------

' '6- carrega os parâmetros Enabled/Disabled dos controles


' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------
' ' - Confirma se o [ Controle ] é [ TrggCtrl ] e [ TrgtCtrl ]
'                        'Confirma se [ sTriggCtrl ] é um [ Trigger ]
'                        ' verifica se o dict [dictTrggCtrls_FilGrp(sForM)] foi criado, o que indica que há [ TrggCtrls ] carregados
'                        vA = IsObject(dictTrggCtrls_FilGrp(sForM))
'
'                        'Se o dicionário de [ TrggCtrls ] não existir ou se ele existir mas [ sCtrL ] não tiver sido incluído, indica que ele NÃO é um trigger
'                        If vA Then vB = dictTrggCtrls_FilGrp(sForM).Exists(sTriggCtrl) Else vB = False
'
'                            If vB Then
'Stop
'                                'Recupera o [ grupo de filtragem ] do [ TriggCtrl ]
'                                If IsObject(dictTrggCtrls_FilGrp(sForM)(sTriggCtrl)) Then
'                                    Set clObjFilGrpsByForm = dictTrggCtrls_FilGrp(sForM)(sTriggCtrl)
'                                    sFilGrp = clObjFilGrpsByForm.sFilGrp
'
'                                End If
'
'                                'Confirma se [ sTriggCtrl ] é também um [ TrgtCtrl ]
'                                ' confirma se o dict [ dictFormFilterGrpsTrgts(sForM) ] existe, o que indica que há Grupos de filtragem no [ Form ]
'                                If IsObject(dictFormFilterGrpsTrgts(sForM)) Then
'
'                                    ' confirma se [ sTriggCtrl ] existe no dict [ dictFormFilterGrpsTrgts(sForM) ], o que indica que ele é um [ TargetCtrl ]
'                                    If dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp) = True Then
'
'                                        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)
'                                        vA = clObjTargtCtrlParam.sTargtCtrlName
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Trigger é Target"
'Stop
'                                    End If
'
'                                End If
'
'                            End If
'
'                        End If
'
'                    End If
' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------






'SpecialEffect
'  fmSpecialEffectFlat
'  fmSpecialEffectRaised * (alto relevo)
'  fmSpecialEffectSunken   (bx relevo)
'  fmSpecialEffectEtched * (esboçado)
'  fmSpecialEffectBump   * (cinzelado)
'
'  * exceto para Checkbox e Option Buttons
'
'
'BackStyle
'  opaque, transp, ...
'
'BorderStyle
'  dashed, solid, ...
'
'ForeColor
'BackColor
'BorderColor
'
'Visible
'Enabled
