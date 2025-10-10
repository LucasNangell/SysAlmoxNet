VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_02aTrggCtrlParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'--------------------------------------------------------
'Classe pra guardar parâmetros de controles de Filtragem
'--------------------------------------------------------

' ccCtrl sQryField sFilGrp iQryFldRmvCharCpt iSrchWildCard iClctdStrSze iSrchOnChange sCascUpDtTrgCtrl
' Trgg.SetorIDfk.Grp01.Rmv03.WdCrd1.Clctd0.SrchOnChange1.CascUpDt(_)---RstAr02

'Public ccCtrl As Control     'Controle
Public sCtrlName As String  'Nome do controle (não é mais necessário pois o nome do controle está como parte do nome do Dicionário)
                            ' nome do controle, usado como Indexador do controle dentro do dicionário

'------------------------------------------------------------
'Propriedades fixas, carregadas na incialização do sistema
'----------------------------------
Public sQryField As String          'Query field do controle pra filtragem
Public sQryFieldCptClean As String  'Query field a ser exibido como título no RecCntCpt (rótulo de contagem de registros)

Public sFilGrp As String    'Grupo de filtragem a que pertence o controle, o que permite identificar  (Grp)
                                ' qual controle deve exibir os resultados da pesquisa e se há outros controles
                                ' a serem considerados

'Public iQryFldRmvCharCpt As String  'Query field limpo do controle pra exibição na contagem de registros  ( Rmv00 )

Public iSrchWildCard As Integer     'Wildcard do controle
                                    ' .se 0   [texto]
                                    ' .se 1   [texto]*
                                    ' .se 2  *[texto]*
'Public iClctdStrSze As Integer      'Indica o comprimento da string do campo calculado que deve ser aproveitada na nova SQL
Public bBolClctd As Boolean         'Indica se o Controle é ou não um campo calculado na Tabela de Dados


Public bOnDirty As Boolean    'indica se o controle deve mudar de cor no sujo
                                ' (se pertencer a um Grupo de Filtragem ou se for um TgtCtrl com indicação de mudança de cor no Sujo)


Public iSrchOnChange As Integer         'Indica se o Trgt Ctrl deve ser atualizado no evento Change ou no LostFocus
                                    ' .se 0   pesquisar durante a digitação
                                    ' .se 1   pesquisar após perder o foco
                                
                                    ' quando o disparador de filtragem for alterado
Public iListboxTxtClmn As Integer  'Indica qual coluna de dados do controle da consulta armazena o valor textual exibido no listbox (usado apenas em Listbox e Combobox)


'Public bMskdCtrlEventFound As Boolean  'a rotina [ MskdTxtbox02_TextMask ] só deve ser ativada se o controle tiver simultaneamente:
'                                       ' . um evento Change com a rotina MskdTxtbox02_TextMask nele incluída
'                                       ' . no [ dictCtrlBehvrParams(sForM) ] o parâmetro [ bMskd ] TRUE
        

Public sCascUpDtTrgCtrl As String   'Indica, caso aplicável, qual controle do Form deve ser atualizado em cascata
                                    
                                    
                                    
'------------------------------------------------------------
'Parâmetro que indica se o controle Trigger deve ou não forçar a atualização do seu Target
'----------------------------------
Public bForceAfterUpdate As Boolean   'a filtragem no [ TrggCtrl ] dispara a atualização do combo [ TrgtCtrl ] e força o evento AfterUpdate pra exibir,
                                      ' nos [ DataFields ] adequados, os dados referentes ao registro selecionado na combo de pré-filtragem
                                      
                                        'revisar estes comentários
                                        '- se o evento não tiver sido chamado manualmente pelo evento do controle na Classe
                                        '  informa ao [ dictColorCtls ] se o Controle deve ou não executar a rotina forçada de Update
                                        '  com alteração de outros controles do form a partir do AfterUpdate deste controle

                                        '- o valor é carregado manualmente pra [ dictColorCtls ] a partir de um dos eventos do próprio controle no Form
                                        '  Se o valor for Verdadeiro e o controle for atualizado chama na Classe [ cls_2aCurrCtlColor ]
                                        '  a rotina que força o evento do controle (vide Form OpsSearch > lstOPs. evento AfterUpdate)


'Public iBdCln As Integer  'Indica qual coluna da consulta tem o valor principal exibido no listbox
'                            '-- SELEÇÃO DE ITENS PREVIAMENTE SELECIONADOS DA LISTBOX --
'                            'se o TargtCtrl tiver o parâmetro [ iBdCln ] maior que ZERO, devem ser incluídos na cláusula WHERE
'                            ' do TriggCtrl associado, os itens previamente selecionados no TrggCtrl
'                            '  .apenas [ TrggCtrls ] associados a [ TargtCtrls ] do tipo [ Listbox ] devem passar por essa avaliação
'                            '  .a cláusula [ WHERE ] do TrggCtrl que irá filtrar o TargtCtrl já foi carregada com os valores preenchidos pelo usuário
'                            '  .a esta cláusula [ WHERE ] deverão ser adicionados os items previamente selecionados no TargtCtrl
'                            '  .para isso a string [ sgbFiLnewFrmWHERE ] deve ser carregada com esses valores adicionais
'                            '
'                            '  .os itens selecionados são guardados em [ ListboxNav.dListboxItems ] à medida que são selecionados
'                            '-----------------------------------------
'                            '-------------------------------------------------------------------------

'------------------------------------------------------------
'Valores dinâmicos obtidos durante o uso do sistema
'----------------------------------
Public dictGetListSrchVals As New Dictionary  'Adiciona no dicionário interno de cada controle, os valores a serem pesquisados
                                              ' apenas em Listbox pode haver mais de um valor. Nos outros controles sempre será apenas um item.
Public dictGetListItemTxts As New Dictionary

Public sGetSQLwhere As String           'String de pesquisa do controle que deve ser acumulada pra formar a condição WHERE do SQL
Public sGetRecCntCptTxt As String          'String a ser exibida no Controle de contagem de registros








