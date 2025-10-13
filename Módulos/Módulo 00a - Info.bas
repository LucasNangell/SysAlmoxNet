Attribute VB_Name = "M�dulo 00a - Info"
Option Compare Database
Option Explicit


'-----------------------------------
'pend�ncias
'-----------------------------------
' Ok -ResetArea

' Ok -Migra��o de eventos de controle pra Classe

' Ok -Dados (Exibi��o)

' Ok - checagem de arquivos de sistema no disco local
'        diferenciar erros de caminho e de aus�ncia de arquivo


' Ok - abrir form Estoque com Edi��o desabilitada. Habilitar assim que um registro for selecionado
' Ok - ajustar EnableDisable pra habilitar [ btnEdtProd ] quando [ lstProdutos ] tiver um registro selecionado

' retirar edi��o de produtos da tela de Estoque. Haver� uma tela exclusiva pra isso.



' -Travar registro na rede --> OpenRecordset(strSQL, dbOpenDynaset, dbPessimistic)

' -Gravar dados (Edi��o, Inclus�o)
' -Gerenciamento de Fotos

' -Servi�os de manuten��o do sistema



'-----------------------------------
'Melhorias
'-----------------------------------
' -Documenta��o
'  . Fluxograma
'  . Roteiro
'  . Par�metros configur�veis
'    . TAGs de controles
'    . Constantes globais

' -Multiselect colocar em uso
' -CascUpDt  colocar em uso
' -Sub forms colocar em uso
' -Log de mudan�a de dados



' dictCtrlType.Add "104", "CommButton"
' dictCtrlType.Add "106", "CheckBox"
' dictCtrlType.Add "107", "OptionGroup"
' dictCtrlType.Add "109", "TextBox"
' dictCtrlType.Add "110", "ListBox"
' dictCtrlType.Add "111", "ComboBox"

'---------------------------------------------------------------------------------------------
' Pontos de parametriza��o do sistema
'---------------------------------------------------------------------------------------------
' . Constantes no m�dulo de Vari�veis

' . Tabelas
'   . Niveis de permiss�o de acesso de usu�rios
'     [ tbl_00(2)bSysUserLoginLevels ]
'     . cadastro de perfis de usu�rio (ao cadastrar novos n�veis de permiss�o lembrar que quanto menor o...)
       
'   . Permiss�es requeridas pra acesso a funcionalidades
'     [ tbl_00(3)aSysRqrdPrmss ]

'   . Par�metros EnblDsbl para controles dos forms
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
' Sequ�ncia de carga de dicion�rios
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
' 1- Garante que os dicion�rios e objetos de classe est�o vazios
'    Call CleanDicts

'---------------------------------------------------------------------------------------------

' '   ---------------------------------------------------------------------------------------------
' '2- Carrega o dicion�rio de CtrlTypes, que ser� usado em trechos do sistema
'     dictCtrlType.Add "104", "CommButton"
' '
' '   ---------------------------------------------------------------------------------------------
' '3- Identifica o usu�rio logado e recupera suas permiss�es de acesso
' '    Set clObjUserParams = New cls_08aLoggedUserParams
'      clObjUserParams.lngUserID = rsTbE.Fields("UserID")
'      ...
'      clObjUserParams.dictUserPermissions.Add CoDtxt, LoginLevelDescri�


' '   ---------------------------------------------------------------------------------------------
' '4- Percorre cada um dos Forms, abre em modo oculto e carrega os dicion�rios do sistema

' '   ---------------------------------------------------------------------------------------------
' '5- pbSub20_TargtCtrlsDictStartUp
        ' pra cada acListBox, acComboBox do [ Form ] avalia se o controle tem a 2a se��o de par�metros
        ' chama [ pbSub21_TargtCtrlsDictBuild ] pra carregar os par�metros e
        ' monta o dicion�rio [ dictFormFilterGrpsTrgts(sForM) ] com os grupos de filtragem do [ Form ]
' '        . inclui no dict o grupo de filtragem associado ao controle ora avaliado
' '        . recupera o SQL do TrgtCtrl associado ao grupo
' '        . checa a inexist�ncia de TrgtCtrls no form e inclui no dict de Log de Erros

' '   ---------------------------------------------------------------------------------------------
' '6- pbSub30_TriggCtrlDictStartUp
        ' se o Form tiver TrgtCtrls passa por todos os acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox do Form
        ' e avalia se o controle tem a TAG necess�ria pra inclus�o nos dicts" & vbCr & "[ dictTrgg00GrpsInForm ] e [ dictTrgg01CtrlsInGrp ]
        ' Avalia se o controle � um dos tipos da categoria BehvrCtrls (acTextBox, acListBox, acComboBox)
        ' Avalia se o controle tem a TAG necess�ria pra inclus�o no dict [ dictCtrlBehvrParams(sForM) ]
        ' e caso postivio inclui o controle no dict
        '
        '-
        '-
'pbSub10_EventsDictBuild
'pbSub31_TriggCtrlDictPreBuild
'pbSub41_CtrlsBehvrDictBuild


' '7- Avalia erros de carga do sistema: controle n�o configurado como [ Trigger] mas com [ bOnDirty ] TRUE


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
' '  1. Durante a carga do sistema se encontrado algum erro � chamada a rotina [ FormStatusBar01_Bld ]
' '   pra montagem do [ dictFormsParams ] com informa��es sobre falhas na carga de par�metros de controles do [ Form ]
' '
' '  2. Ao final da carga do sistema se o par�metro [ bForm_ShowWarns ] do dict [ dictFormsParams ] for TRUE
' '     chama a rotina [ FormStatusBar02_SetWarn(fForM) ] pra alterar a Status Bar do Form rec�m carregado
' '     e exibir o alerta de erros na carga do sistema
' '
' '  3. A rotina [ FormLoad06b_BackFromFormLoad(sForM) ] � chamada quando for necess�rio avaliar se [ bForm_ShowWarns ] � TRUE

' '  4. Ap�s a abertura do Form se a [ StatusBar ] receber um duplo clique
' '     � chamada a rotina [ FormStatusBar04_OpnLogForm ] pra montar o texto a ser exibido
' '     no Form [ frm_00(1)cSysLoadLog ] indicando o Log de alertas de carga do sistema
' '
' '
' '
' '
' '

' '
' '  FormStatusBar01_Bld(sForM, "...", sLoadLogWarn)
' '  ... MissingBehvrParams
' '   "H� controles no Formul�rio sem [ BehvrParams ]. Esses controles n�o ir�o se comportar como esperado."
' '
' '  ... NoTrgtCtrls
' '   "N�o foram encontrados TargetCtrls no formul�rio. N�o ser� poss�vel fazer pesquisas."
' '
' '  ... DirtyTrue_NoTrgg
' '   "H� controles n�o carregados como [ Trigger ] mas configurados com [ OnDirty ]"
                
                
                
' '  ----------------------------------------------------------------
' '  Roteiro pra disparo de pesquisas em [ TriggCtrls ]
' '  ---------------------------------------------------------------------------------------------
' '  1. Disparado o evento [ Change ] do pr�prio controle
' '
' '  2. Chama a rotina [ MskdTxtbox02_TextMask ] pra tratar o campo
' '     de forma a aplicar a m�scara definida pro controle
' '

' '     substitu�do por [ bMskdCtrlEventFound ] apenas dentro da rogina BuildWhere
' '     Atribui TRUE � vari�vel [ bTriggrddByCtrlEvent ] pra garantir que
' '     o c�digo de pesquisa desconsidere a m�scara do controle

' '  3. Segue pro evento Change da Classe [ cls_10aCtrls_Events ] de eventos de controle
' '  4. Chama a rotina [ pb_TargtCtrlUpdate00_TimerDelay ] pra iniciar o timer de atualiza��o do controle
' '  5. Chama a rotina [ pb_TargtCtrlUpdate01_Start ] pra iniciar a atualiza��o do controle
' '  6. Chama a rotina [ pb_TargtCtrlUpdate03_UNIQUEupdate ] pra identificar:
' '      - o Grupo de Filtragem do controle disparador
' '      - o tipo de atualiza��o disparada:
' '        . Reset Unique
' '        . Reset Area
' '  8. Chama a rotina [ pb_TargtCtrlUpdate06_BuildWHERE ] pra iniciar a montagem da cl�usula WHERE pra atualiza��o do [ TrgtCtrl ]
' '       Se o controle n�o tiver sido disparado pelo pr�prio evento [ Change ] trata o controle
' '       como se fosse [ Masked ] FALSE
' '  9. Chama a rotina [ BuildSQL (_TextBox) ] apropriada a depender do tipo de controle disparador
' '      se a configura��o [ Masked ] do controle for TRUE, significa al�m de ter TRUE no par�metro
' '      a pesquisa foi disparada pelo evento [ Change ] do pr�prio controle, o que confirma que ele � um [ Masked ]
' '  ----------------------------------------------------------------
                
                

' '  ----------------------------------------------------------------
' '  Roteiro de montagem de dicion�rios
' '  ---------------------------------------------------------------------------------------------
' '
' '   ---------------------------------------------------------------------------------------------
' '   Percorre cada um dos Forms, abre em modo oculto e carrega os dicion�rios do sistema
' '    em seguida fecha o Form (os objetos carregados nas classes s�o perdidos ao fechar o Form
' '    ent�o descartei o carregamento direto de objetos)
    
' '1- carrega os dados dos TargtCtrls (Listbox e Combobox)
' '  Tratamentos de erro (CHECAGENS)
' '   - Par�metro RecCntCtrl vazio
' '      . Ao executar uma filtragem o sistema n�o tenta exibir a contagem de registros
' '        se o par�metro for vazio
' '
' '   - RecCntCtrl associado ao TargtCtrl n�o existente no form
' '      . Exibe alerta, continua a carga do controle mas deixa o par�metro RecCntCtrl vazio.
' '        Assim evita que o sistema tente encontr�-lo pra exibir a contagem de registros ap�s filtragem.
' '
' '   - Mais de um TargtCtrl associado ao mesmo Grupo de Filtragem
' '      . D� alerta e n�o carrega o Listbox como TargtCtrl
' '
' '. pbSub20_TargtCtrlsDictStartUp
' ' . pbSub21_TargtCtrlsDictBuild
' '    Cria uma nova varia��o do dicion�rio pro Formul�rio corrente
'       Set dictFormFilterGrpsTrgts(sForm) = New Dictionary
'       Set dDict = dictFormFilterGrpsTrgts(sForm)
' '
' '   Recupera o Grupo de Filtragem pra montar o Dicion�rio
' '   . pbSub06_GetCtrlTagFltrGrp
' '
' '   Cria um novo objeto [ clObjTargtCtrlParam ] da Classe [ cls_01aTargtCtrlParams_Evnts ]
' '    pra ser inclu�do no Dict [ dictFormFilterGrpsTrgts(sForm) ]
'      Set clObjTargtCtrlParam = New cls_01aTargtCtrlParams_Evnts
' '
' '   Adiciona um novo item no dicion�rio [ dictFormFilterGrpsTrgts ] e guarda nele o objeto [ clObjTargtCtrlParam ]
' '    com os respectivos par�metros do targtCtrl definidos na classe [ cls_01aTargtCtrlParams_Evnts ]
'       dDict.Add sFilGrp, clObjTargtCtrlParam
' '
' '    Obs: n�o foi prevista a possibilidade de incluir num Grupo de Filtragem mais de um TargtCtrl
' '          pra isso funcionar, ao inv�s de incluir no [ dictFormFilterGrpsTrgts(sForm) ] um objeto de classe
' '          com os par�metros do Grupo (TargtCtrl, Where, etc), teria que ser inclu�do um segundo dicion�rio.
' '          Nesse segundo dicion�rio seriam inclu�dos cada um dos TargtCtrls associados ao Grupo, e nesse dicion�rio
' '          a� sim seria inclu�do o [ clObjTargtCtrlParam ]
' '
' '     Atribui ao Listbox os par�metros esperados pela Classe [ cls_01aTargtCtrlParams_Evnts ]
' '     -------------------------------------------------------------------------------------------
' '      dictFormFilterGrpsTrgts   - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictFormFilterGrpsTrgts(sForm)     - Grupos de Filtragem
' '          Item "01"  >  clObjTargtCtrlParam  - par�metros do Grupo de Filtragem
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

' '   Recupera o Grupo de Filtragem pra montar o Dicion�rio
' '   . pbSub06_GetCtrlTagFltrGrp

' '   Cria o dicion�rio com grupos de filtragem pra cada formul�rio
'     Set dictTrgg00GrpsInForm(sForm) = New Dictionary

' '   Cria o nested Dict com os controles associados a cada Grupo de Filtragem
'      Set dictTrgg01CtrlsInGrp(sFilGrp) = New Dictionary
'      Set dDict = dictTrgg01CtrlsInGrp(sFilGrp)
 
' '   Cria um novo objeto [ clObjTriggCtrlParam ] da Classe [ cls_02aTrggCtrlParams ]
' '    pra ser inclu�do no Dict [ dictTrgg01CtrlsInGrp(sFilGrp) ]
'       Set clObjTriggCtrlParam = New cls_02aTrggCtrlParams
'       dDict.Add sCtrL, clObjTriggCtrlParam
' '   ---------------------------------------------------------------------------------------------
' '      dictTrgg00GrpsInForm  - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictTrgg00GrpsInForm(sForM)     - Grupos de Filtragem
' '          Item "01"   > dictTrgg01CtrlsInGrp(sFilGrp) - Controles associados ao Grupo
' '                          Item "cmbSrcCateg"  >  clObjTriggCtrlParam  - par�metros do Controle
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

 
' '2b- Cria o dicion�rio com os Trigg Controls de cada formul�rio, indicando o Grupo de Filtragem associado a cada controle
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
' '          Item "cmbSrcCateg"  > clObjFilGrpsByForm  - par�metros do Controle
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

' '   Recupera a ID da ResetArea pra montar o Dicion�rio

' '   Cria o dicion�rio com Reset Areas pra cada formul�rio
'     Set dictFrmResetAreas(sForM) = New Dictionary
'     Set dDictOuter = dictFrmResetAreas(sForM)

' '    Cria o nested Dict com os controles associados a cada ResetArea
'      Set dictResetAreaCtrls(sRstArea)= New Dictionary
'      Set dDictInner = dictResetAreaCtrls(sRstArea)
 
' '    Cria um novo objeto [ clObjRstAreaParams ] da Classe [ cls_05aResetAreasParams ]
' '     pra ser inclu�do no Dict [ dictResetAreaCtrls(sRstArea) ]
'      Set clObjRstAreaParams = New cls_05aResetAreasParams
'      dDictInner.Add sCtrl, clObjRstAreaParams
' '   ---------------------------------------------------------------------------------------------
' '      dictFrmResetAreas  - Forms
' '        Item "frm_01(1)bProdEstoque"
' '
' '        dictFrmResetAreas(sForM)     - �reas de Reset
' '          Item "01"   > dictResetAreaCtrls(sRstArea) - Controles associados � �rea de Reset
' '                          Item "cmbSrcCateg"  >  clObjRstAreaParams  - par�metros do Controle
' '                                                  .sQryField = sQryField
' '                                                  .sRstAr = sRstAr
' '   ---------------------------------------------------------------------------------------------

' '6- carrega os par�metros Enabled/Disabled dos controles


' '   ---------------------------------------------------------------------------------------------
' '   ---------------------------------------------------------------------------------------------
' ' - Confirma se o [ Controle ] � [ TrggCtrl ] e [ TrgtCtrl ]
'                        'Confirma se [ sTriggCtrl ] � um [ Trigger ]
'                        ' verifica se o dict [dictTrggCtrls_FilGrp(sForM)] foi criado, o que indica que h� [ TrggCtrls ] carregados
'                        vA = IsObject(dictTrggCtrls_FilGrp(sForM))
'
'                        'Se o dicion�rio de [ TrggCtrls ] n�o existir ou se ele existir mas [ sCtrL ] n�o tiver sido inclu�do, indica que ele N�O � um trigger
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
'                                'Confirma se [ sTriggCtrl ] � tamb�m um [ TrgtCtrl ]
'                                ' confirma se o dict [ dictFormFilterGrpsTrgts(sForM) ] existe, o que indica que h� Grupos de filtragem no [ Form ]
'                                If IsObject(dictFormFilterGrpsTrgts(sForM)) Then
'
'                                    ' confirma se [ sTriggCtrl ] existe no dict [ dictFormFilterGrpsTrgts(sForM) ], o que indica que ele � um [ TargetCtrl ]
'                                    If dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp) = True Then
'
'                                        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)
'                                        vA = clObjTargtCtrlParam.sTargtCtrlName
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Trigger � Target"
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
'  fmSpecialEffectEtched * (esbo�ado)
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
