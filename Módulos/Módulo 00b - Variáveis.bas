Attribute VB_Name = "Módulo 00b - Variáveis"
Option Compare Database
Option Explicit

'Sleep function
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

''--- CONSTANTES PUBLICAS -----------------------
'--  de Sistema ---------------------------------
      Public Const gBiTypingDelay  As Integer = 250                        'tempo de espera entre a digitação nos campos de filtragem e o disparo da filtragem
      Public Const gBsSystemDefaultForm As String = "frm_00(1)aSysStart"   'nome do formulário de início de sistema
      
      'Ativa/Inativa tratamentos de erro
      Public Const gBbEnableErrorHandler As Boolean = False
      
      
      'Carga do sistema: rotina [ SysLoad01_SysDictsLoad ], nível 1
      Public Const gBbDepurandoLv01a As Boolean = False 'false  true
      
      'Subrotinas da rotina principal, nível 2
      Public Const gBbDepurandoLv01b As Boolean = False
      
      'Subrotinas das rotinas secundárias, nível 3
      Public Const gBbDepurandoLv01c As Boolean = False
      
      
      Public Const gBbDepurandoLv02a As Boolean = False   'Abertura de Forms
      Public Const gBbDepurandoLv03a As Boolean = False   'Filtragem
      
      Public Const gBbDebugOn As Boolean = False
      
      Public Const gBbInitCtrlEvents As Boolean = True
      
      Public Const bBsStatusBarLabel As String = "lblStatusBar"
      
      Public Const GbLngTrgtHorizOffset As Long = 0 '300
      Public Const GbLngTrgtVertOffset As Long = 0 ' 400
      
      Public Const GbLngTrgtLeft As Long = 500
      Public Const GbLngTrgtTop As Long = 700
      
      
    '----------------------------------------------
    '--  de caminhos de arquivos ---------------------------------
      'Public Const GBsSysLocalIconFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administração\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\SysIcons\"
            '---------
      
      'Constantes Globais com a raiz dos caminhos locais do sistema
      '-----------------------------------------------------------
      ' ... Almox - Sistema\
      Public Const gBsSysLocalFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administração\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\"
      ' ... Almox - Sistema\SysIcons\
      Public Const gBsSysLocalIconFolder As String = gBsSysLocalFolder & "SysIcons\"
      ' ... Almox - Sistema\UpdateFiles\
'parei aqui2: após a finalização do sistema retirar o comentário da linha abaixo
      'Public Const gBsSysLocalUpdateFolder As String = gBsSysLocalFolder & "UpdateFiles\"
      
      
      'Constantes Globais com os icones usados no sistema
      '-----------------------------------------------------------
      Public Const gBsSysLocalIcoTrggUnlock As String = gBsSysLocalIconFolder & "frmDock_UnlockToTrgg24.png"
      Public Const gBsSysLocalIcoTrggLock As String = gBsSysLocalIconFolder & "frmDock_LockToTrgg24.png"
      Public Const gBsSysLocalIcoTrgtUnlock As String = gBsSysLocalIconFolder & "frmDock_UnlockToTrgt24.png"
      Public Const gBsSysLocalIcoTrgtLock As String = gBsSysLocalIconFolder & "frmDock_LockToTrgt24.png"
      
      Public Const gBsSysLocalIcoAtoZ As String = gBsSysLocalIconFolder & "frmOPsSearch-Sort - AtoZ.png"
      Public Const gBsSysLocalIcoZtoA As String = gBsSysLocalIconFolder & "frmOPsSearch-Sort - ZtoA.png"
      
      'global com o a raiz do caminho do sistema na Rede
      '... BaseDeDados\
      Public Const gBsSysNetFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administração\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\"
      '... BaseDeDados\Imagens
      Public Const gBsSysNetImgsFolder As String = gBsSysNetFolder & "Imagens\"
      '... BaseDeDados\ClientVersionUpdate
      
'parei aqui2: após a finalização do sistema retirar o comentário da linha abaixo
      'Public Const gBsSysNetClientVrsUpdFolder As String = gBsSysNetFolder & "ClientsVersionUpdate\"
    '----------------------------------------------
      
'----------------------------------------------
'-----------------------------------
'--  de Cor de Controles (cor PRETO, preenchida em Propriedades #000000; na roda de cores RGB(0,0,0); em código 0)
      
      Public Const GbLngSTATUSbarAlert As Long = 1128094  'Vermelho escuro
      Public Const GbLngSTATUSbarNoAlert As Long = 4934479  'Cinza escuro

      'Cores para botões (desabilitados)
      Public Const GbLngBtnGREyBackColor As Long = 12566463
      Public Const GbLngBtnGREyBdColor As Long = 8355711
      'Public Const sgbBtnGREyBackColor As String = "#ADABAB"
      'Public Const sgbBtnGREyBdColor As String = "#969696"

      'Cores para demais controles (desabilitados), exceto Checkbox e Optionbutton
      Public Const GbLngTxtBaseBackColor As Long = 9609156
      Public Const GbLngTxtBASeForeColor As Long = 6118789
      'Public Const lngGbBtnBaseBdColor As Long = ??

      Public Const GbLngHLclrBackColor As Long = 9691901 'Amarelo
      Public Const GbLngDIRTclrBackColor As Long = 13487516  'Verde


'-----------------------------------
'----------------------------------------------

''--- VARIAVEIS PUBLICAS -----------------------
'
''--------------------
''--  System parameters ----------
      'Public sGbQrySQLfrmProdSupEstoque As String
      Public gBbDictsLoaded As Boolean   'Indica que foi feita a carga completa do sistema
      'Public gBbNoTrgtCtrls As Boolean   'Indica se foi criado pelo menos um dicionário de [ TrgtCtrls ]
      Public gBbTrgtCtrlsFound As Boolean   'Indica se foi criado pelo menos um dicionário de [ TrgtCtrls ]
      
'      Public sGbFiLnewFrmSELECT As String      'RowSource original do ListBox ou Combo ao carregar um Form, 1a parte (do Select)
'      Public sGbFiLnewFrmWHERE As String       'RowSource original do ListBox ou Combo, ao carregar um Form, 2a parte (do Where)
'      Public sGbFiLnewFrmOrderBy As String     'RowSource original do ListBox ou Combo, ao carregar um Form, 3a parte (do Order)
'      Public sGbFiLnewFrmTmpWHERE As String    'condição WHERE atual do controle, que pode ser diferente da original



''--  de Rotinas de auxiliares
'      Public gBbCallingByName As Boolean   'usada pra confirmar se o Controle tem o eventoIndica se foi criado pelo menos um dicionário de [ TrgtCtrls ]
'      Public gBbEventFound As Boolean   'retorna se o evento pesquisado existe
      
''--  de Rotinas de Filtragem
        'variáveis globais temporárias pra guardar o Controle e o Form necessários
        ' pra chamar a rotina de atraso de filtragem durante digitação
      Public gBcTrggCtrl As Control
      Public gBfTrggCtrlForm As Form


      
'--------------------------------
'--------------------------------
''--  Dicionários e Classes ----
''------------------------------
'--------------------------------

       'Dicionários pra guardar os TriggCtrls do sistema,
       ' agrupados por Formulário e por Grupo de filtragem
       
       ''Removido pois o parâmetro foi incluído no Dict [ dictFormsParams ]
       'Public dictFormStatusText As New Dictionary   'dicionário pra guardar eventuais mensagens de Status a serem exibidas quando o form for aberto
     '----------------
       Public dictFndFldrVars As New Dictionary     'dicionário de variáveis de Paths
     '----------------
     
     '----------------
       Public dictTempDict As New Dictionary     'dicionário temporário pra ser usado localmente em vários pontos do sistema
     '----------------
       
       
     '----------------
       Public dictCtrlTypeStR As New Dictionary     'dicionário de tipos de variável, com a String do nome
       Public dictCtrlTypeShort As New Dictionary   'dicionário de tipos de variável, com a abreviatura
     '----------------
       
     '----------------
     ' dict pra guardar as permissões de acesso do usuário logado
      Public dictUserPermissions As New Dictionary

     ' objeto de classe pra armazenar parâmetros dos Controles pra alternar o Status Enable/Disable
       Public clObjUserParams As cls_08aLoggedUserParams   'declaração do objeto de classe
     '----------------
       
       
       
     '----------------
     ' dict pra guardar os [ Grupos de filtragem ] do formulário
     '  e os respectivos controles do Grupo
       Public dictFormFilGrpsEnDsAllCtrls As New Dictionary
       
     ' dict pra guardar os [ Grupos de filtragem ] do formulário
     '  e o respectivo [ TrgtCtrl ] na classe
       Public dictFormFilterGrpsTrgts As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os parâmetros
    
     ' objeto de classe pra armazenar parâmetros dos Listbox dos formulários
       Public clObjTargtCtrlParam As cls_01aTargtCtrlParams_Evnts   'declaração do objeto de classe
     
     ' dict pra guardar os Listboxes do sistema
     '  e o respectivo [ Grupo de Filtragem ]
       Public dictTrgtCtrlsFilterGrps As New Dictionary
     '----------------
       
       
     '----------------
       Public dictTrgg00GrpsInForm As New Dictionary  'dicionário com grupos de filtragem para o formulário
       Public dictTrgg01CtrlsInGrp As New Dictionary  'dicionário com os controles associados a cada grupo de filtragem
       
     ' objeto de classe pra armazenar parâmetros dos TriggCtrls
       Public clObjTriggCtrlParam As cls_02aTrggCtrlParams   'declaração do objeto de classe
     '----------------
       
     '----------------
     ' dict pra guardar os [ TriggCtrls ] de cada Formulário indicando o Grupo de Filtragem associado
       Public dictTrggCtrlsInForm As New Dictionary

     ' objeto de classe pra armazenar o Grupo de Filtragem associado aos controles de cada Formulário
       Public clObjFilGrpsByForm As cls_03aCtrlsGrpsByForm   'declaração do objeto de classe
     '----------------
       
       
       
       
     '----------------
     ' dict pra guardar as áreas de Reset, por Formulário e dentro da Área os controles pertencentes
       Public dictFrmResetAreas As New Dictionary
       'Public dictResetAreaCtrls As New Dictionary
            'key   > nro sequencial da Area de Reset
            'valor > nome do controle
      
     '  objeto de classe pra armazenar parâmetros dos Controles pertencentes a áreas de Reset
        Public clObjRstAreaParams As cls_05aResetAreasParams   'declaração do objeto de classe
      
     ' dict pra guardar os botões de Reset do Formulário
       Public dictRstArBTNsByNr As New Dictionary        'dicionário pra guardar os botões de Reset do form, indexando pela Area de Reset
       Public dictRstArBTNsByName As New Dictionary      'dicionário pra guardar os botões de Reset do form, indexando pelo nome do controle
     '----------------
      
     '----------------
     ' dict pra cada Form pra guardar o objeto de classe com parâmetros dos Controles
     '  que vão alternar entre Enable/Disable
       Public dictCtrlEnblDsblParams As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os parâmetros
     
     ' objeto de classe pra armazenar parâmetros dos Controles
       Public clObjCtrlsEnblDsblParams As cls_07aCtrlsEnblDsblParams   'declaração do objeto de classe
             
     ' -->  dicionário, dentro do objeto de classe pra armazenar os parâmetros separados por Enabled e Disabled
            Public dictParamByLckdStatus As New Dictionary

     ' --->    objeto de classe pra armazenar parâmetros dos Controles
               Public clObjLckdStatusParam As cls_07bLckdStatusParams   'declaração do objeto de classe
     '----------------
       
     '----------------
     ' objeto de classe para armazenar parâmetros utilizados nas rotinas de abertura de formulário
       Public clObjFormOpenParams As cls_09cParamsToOpenForms
       
     '----------------
     ' dict pra guardar parâmetros de comportamento de Triggers e outros controles, por form
       Public dictCtrlBehvrParams As New Dictionary
     
     ' objeto de classe pra armazenar parâmetros dos Listbox dos formulários
       Public clObjCtrlBehvrParams As cls_11aCtrlBehvrParams   'declaração do objeto de classe
     '----------------

     '----------------
     ' dict pra guardar parâmetros de Formulários
       Public dictFormsParams As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os parâmetros
     
     ' objeto de classe pra armazenar parâmetros dos formulários
       Public clObjFormsParams As cls_09aFormsParams   'declaração do objeto de classe
     
' ----> objeto de classe interno à Classe [ clObjFormsParams ] pra armazenar alertas de erro de carga de formulários
        Public clObjStatusBarWarn As cls_09bFormsParamsLoadLogTxt   'declaração do objeto de classe
     '----------------


     '----------------
     ' dict pra carregar Controles cujos eventos devem ser monitorados
       Public dictCtrlsEvents As New Dictionary
     
     ' objeto de classe pra armazenar parâmetros dos Listbox dos formulários
       Public clObjCtrlsEvents As cls_10aCtrls_Events   'declaração do objeto de classe
     '----------------


     '----------------
     ' dict pra guardar os DataFields dos controles, por form
       Public dictFormDataFlds00Ctrls As New Dictionary   'dicionário com os controles tipo DataField do [ form ]
       Public dictFormDataFlds01Grps As New Dictionary   'dicionário com os Grupos associados aos controles tipo DataField
     
     
     ' objeto de classe pra armazenar parâmetros dos Listbox dos formulários
       Public clObjCtrlDataFieds As cls_04aCtrlsDataFields
     '----------------

      
     '----------------
      'dict para guardar as consultas originais dos controles [ Listbox ] e [ Combobox ]
      Public dictFormQrysCtrls As New Dictionary
     '----------------


     '----------------
     ' dict pra guardar os DataFields dos controles, por form
       Public dictFormCommButtons As New Dictionary
     
     ' objeto de classe pra armazenar parâmetros dos Listbox dos formulários
       Public clObjCommButtons As cls_12aCommButtonParams
     '----------------




'----------------------------------------------
'--  de Cor ----
'    'Types: declaração de variáveis para a função LongDecToRGB_HEX_HSL
      Type MyColors
         CorAccessGBlng As Long
         CorRGBgbStr As String
         CorHEXgbStr As String
         CorHSLgbStr As String
      End Type
      Public CoRes As MyColors
    
'-   'Types: declaração de variáveis para a função que identifica o RGB da cor from DEC
      Type RGBcolor
         Red As Long
         Green As Long
         Blue As Long
      End Type
      Public GetRGB As RGBcolor

'-   'Types: declaração de variáveis para a função que identifica o valor HSL
      Type HSLcolor
         lHue As Long
         lSat As Long
         lLum As Long
         sTintButton As String
      End Type
      Public GetHSL As HSLcolor


'----------------------------------------------
'--  de ... ----
     'Types: declaração de variáveis pra armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
      Type sLstbxSQLstr
          sLstbxSQL_aSELECT As String
          sLstbxSQL_bFROM As String
          sLstbxSQL_cWHERE As String
          sLstbxSQL_dOrderBy As String
          sLstbxSQL_eMAIN As String    'sLstbxSQL_aSELECT & " " & sLstbxSQL_bFROM
          
      End Type
      Public sGbQrySQLstr As sLstbxSQLstr
'-----------------------------------
'----------------------------------------------

'--  de ... ----
     'Types: declaração de variáveis pra armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
      Type vCtrlPrmissGrnted
          bPermissionGrated As Boolean
          sCtrlNewTipText As String
          
      End Type
      Public GetPrmissGrntedType As vCtrlPrmissGrnted
'-----------------------------------
'----------------------------------------------
      
'--  de ... ----
     'Types: declaração de variáveis pra
      Type vCheckQryFld
          bFoundQryFld As Boolean     'Verdadeiro se o campo ora analisado estiver presente no Grid da consulta
          sQry As String              'usado apenas pra recuperar o nome consulta ora analisada, como String pra ser exibido na msg de erro
          
      End Type
      Public NstdVarQryFld As vCheckQryFld

'-----------------------------------
'----------------------------------------------
      
''--  de ... ----
'     'Types: declaração de variáveis pra
'      Type lngFormCoordinates
'          lngTrgtHorizPos As Long
'          lngTrgtVertPos As Long
'
'
'      End Type
'      Public NstdFrmCoords As lngFormCoordinates
'''-----------------------------------
'''----------------------------------------------
      


' ---------------------------------------------------------------
' --  Modelo ----
'     'Types: declaração de variáveis para armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
'     ' User Defined Types
'      Type vUserType
'          sCtrlWhere As String            'string SQL WHERE com o texto que deve ser usado pra pesquisa dos dados
'          sCtrlReCntCption As String     'string para exibição na contagem de registros
'      End Type
'      Public NstdVar As vUserType
'
'
'       vUserType
'       NstdVar             NstdVar = GetFunction(var01 as ...
'       GetFunction         Public Function GetSQL_01_Chkbox(iFrmIndexID As Integer, ...


' ---------------------------------------------------------------
'      'chamada da função, na rotina principal, pra obtenção dos valores
'      '  NstdVar = GetFunction(iFrmIndexID, cTriggCtrl, ...)
'      '
' ---------------------------------------------------------------
'      ' função pra obtenção dos valores
'      '   Public Function GetFunction(iFrmIndexID As Integer, cTriggCtrl As Control, ...) As vUserType
'      '     GetFunction.sCtrlWhere = sWhere
'      '     GetFunction.sCtrlReCntCption = sReCntCptn
' ---------------------------------------------------------------

' ---------------------------------------------------------------
'      '
'      'retorno à chamada, usando os valores
'      '   vA = NstdVar.sCtrlWhere
'      '   vB = NstdVar.sCtrlReCntCption
' --  Modelo ----
' ---------------------------------------------------------------


                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
''
'' A verificar se deve ser levado de volta pra área de Declaração de variáveis

'#If VBA7 Then
'Private Type ChooseColor
'    lStructSize               As Long
'    hwndOwner                 As LongPtr
'    hInstance                 As LongPtr
'    rgbResult                 As Long
'    lpCustColors              As LongPtr
'    flags                     As Long
'    lCustData                 As LongPtr
'    lpfnHook                  As LongPtr
'    lpTemplateName            As String
'End Type
'
'#Else
'
'Private Type ChooseColor
'    lStructSize               As Long
'    hwndOwner                 As Long
'    hInstance                 As Long
'    rgbResult                 As Long
'    lpCustColors              As Long
'    flags                     As Long
'    lCustData                 As Long
'    lpfnHook                  As Long
'    lpTemplateName            As String
'End Type
'
'#End If

''
''
'
'
'
'
