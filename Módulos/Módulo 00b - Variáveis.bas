Attribute VB_Name = "M�dulo 00b - Vari�veis"
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
      Public Const gBiTypingDelay  As Integer = 250                        'tempo de espera entre a digita��o nos campos de filtragem e o disparo da filtragem
      Public Const gBsSystemDefaultForm As String = "frm_00(1)aSysStart"   'nome do formul�rio de in�cio de sistema
      
      'Ativa/Inativa tratamentos de erro
      Public Const gBbEnableErrorHandler As Boolean = False
      
      
      'Carga do sistema: rotina [ SysLoad01_SysDictsLoad ], n�vel 1
      Public Const gBbDepurandoLv01a As Boolean = False 'false  true
      
      'Subrotinas da rotina principal, n�vel 2
      Public Const gBbDepurandoLv01b As Boolean = False
      
      'Subrotinas das rotinas secund�rias, n�vel 3
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
      'Public Const GBsSysLocalIconFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administra��o\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\SysIcons\"
            '---------
      
      'Constantes Globais com a raiz dos caminhos locais do sistema
      '-----------------------------------------------------------
      ' ... Almox - Sistema\
      Public Const gBsSysLocalFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administra��o\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\"
      ' ... Almox - Sistema\SysIcons\
      Public Const gBsSysLocalIconFolder As String = gBsSysLocalFolder & "SysIcons\"
      ' ... Almox - Sistema\UpdateFiles\
'parei aqui2: ap�s a finaliza��o do sistema retirar o coment�rio da linha abaixo
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
      Public Const gBsSysNetFolder As String = "\\redecamara\dfsdata\CGraf\Sefoc\Administra��o\Projetos Diversos Sefoc\Melhoria de Processos CGraf\Almox - Sistema\BaseDeDados\"
      '... BaseDeDados\Imagens
      Public Const gBsSysNetImgsFolder As String = gBsSysNetFolder & "Imagens\"
      '... BaseDeDados\ClientVersionUpdate
      
'parei aqui2: ap�s a finaliza��o do sistema retirar o coment�rio da linha abaixo
      'Public Const gBsSysNetClientVrsUpdFolder As String = gBsSysNetFolder & "ClientsVersionUpdate\"
    '----------------------------------------------
      
'----------------------------------------------
'-----------------------------------
'--  de Cor de Controles (cor PRETO, preenchida em Propriedades #000000; na roda de cores RGB(0,0,0); em c�digo 0)
      
      Public Const GbLngSTATUSbarAlert As Long = 1128094  'Vermelho escuro
      Public Const GbLngSTATUSbarNoAlert As Long = 4934479  'Cinza escuro

      'Cores para bot�es (desabilitados)
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
      'Public gBbNoTrgtCtrls As Boolean   'Indica se foi criado pelo menos um dicion�rio de [ TrgtCtrls ]
      Public gBbTrgtCtrlsFound As Boolean   'Indica se foi criado pelo menos um dicion�rio de [ TrgtCtrls ]
      
'      Public sGbFiLnewFrmSELECT As String      'RowSource original do ListBox ou Combo ao carregar um Form, 1a parte (do Select)
'      Public sGbFiLnewFrmWHERE As String       'RowSource original do ListBox ou Combo, ao carregar um Form, 2a parte (do Where)
'      Public sGbFiLnewFrmOrderBy As String     'RowSource original do ListBox ou Combo, ao carregar um Form, 3a parte (do Order)
'      Public sGbFiLnewFrmTmpWHERE As String    'condi��o WHERE atual do controle, que pode ser diferente da original



''--  de Rotinas de auxiliares
'      Public gBbCallingByName As Boolean   'usada pra confirmar se o Controle tem o eventoIndica se foi criado pelo menos um dicion�rio de [ TrgtCtrls ]
'      Public gBbEventFound As Boolean   'retorna se o evento pesquisado existe
      
''--  de Rotinas de Filtragem
        'vari�veis globais tempor�rias pra guardar o Controle e o Form necess�rios
        ' pra chamar a rotina de atraso de filtragem durante digita��o
      Public gBcTrggCtrl As Control
      Public gBfTrggCtrlForm As Form


      
'--------------------------------
'--------------------------------
''--  Dicion�rios e Classes ----
''------------------------------
'--------------------------------

       'Dicion�rios pra guardar os TriggCtrls do sistema,
       ' agrupados por Formul�rio e por Grupo de filtragem
       
       ''Removido pois o par�metro foi inclu�do no Dict [ dictFormsParams ]
       'Public dictFormStatusText As New Dictionary   'dicion�rio pra guardar eventuais mensagens de Status a serem exibidas quando o form for aberto
     '----------------
       Public dictFndFldrVars As New Dictionary     'dicion�rio de vari�veis de Paths
     '----------------
     
     '----------------
       Public dictTempDict As New Dictionary     'dicion�rio tempor�rio pra ser usado localmente em v�rios pontos do sistema
     '----------------
       
       
     '----------------
       Public dictCtrlTypeStR As New Dictionary     'dicion�rio de tipos de vari�vel, com a String do nome
       Public dictCtrlTypeShort As New Dictionary   'dicion�rio de tipos de vari�vel, com a abreviatura
     '----------------
       
     '----------------
     ' dict pra guardar as permiss�es de acesso do usu�rio logado
      Public dictUserPermissions As New Dictionary

     ' objeto de classe pra armazenar par�metros dos Controles pra alternar o Status Enable/Disable
       Public clObjUserParams As cls_08aLoggedUserParams   'declara��o do objeto de classe
     '----------------
       
       
       
     '----------------
     ' dict pra guardar os [ Grupos de filtragem ] do formul�rio
     '  e os respectivos controles do Grupo
       Public dictFormFilGrpsEnDsAllCtrls As New Dictionary
       
     ' dict pra guardar os [ Grupos de filtragem ] do formul�rio
     '  e o respectivo [ TrgtCtrl ] na classe
       Public dictFormFilterGrpsTrgts As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os par�metros
    
     ' objeto de classe pra armazenar par�metros dos Listbox dos formul�rios
       Public clObjTargtCtrlParam As cls_01aTargtCtrlParams_Evnts   'declara��o do objeto de classe
     
     ' dict pra guardar os Listboxes do sistema
     '  e o respectivo [ Grupo de Filtragem ]
       Public dictTrgtCtrlsFilterGrps As New Dictionary
     '----------------
       
       
     '----------------
       Public dictTrgg00GrpsInForm As New Dictionary  'dicion�rio com grupos de filtragem para o formul�rio
       Public dictTrgg01CtrlsInGrp As New Dictionary  'dicion�rio com os controles associados a cada grupo de filtragem
       
     ' objeto de classe pra armazenar par�metros dos TriggCtrls
       Public clObjTriggCtrlParam As cls_02aTrggCtrlParams   'declara��o do objeto de classe
     '----------------
       
     '----------------
     ' dict pra guardar os [ TriggCtrls ] de cada Formul�rio indicando o Grupo de Filtragem associado
       Public dictTrggCtrlsInForm As New Dictionary

     ' objeto de classe pra armazenar o Grupo de Filtragem associado aos controles de cada Formul�rio
       Public clObjFilGrpsByForm As cls_03aCtrlsGrpsByForm   'declara��o do objeto de classe
     '----------------
       
       
       
       
     '----------------
     ' dict pra guardar as �reas de Reset, por Formul�rio e dentro da �rea os controles pertencentes
       Public dictFrmResetAreas As New Dictionary
       'Public dictResetAreaCtrls As New Dictionary
            'key   > nro sequencial da Area de Reset
            'valor > nome do controle
      
     '  objeto de classe pra armazenar par�metros dos Controles pertencentes a �reas de Reset
        Public clObjRstAreaParams As cls_05aResetAreasParams   'declara��o do objeto de classe
      
     ' dict pra guardar os bot�es de Reset do Formul�rio
       Public dictRstArBTNsByNr As New Dictionary        'dicion�rio pra guardar os bot�es de Reset do form, indexando pela Area de Reset
       Public dictRstArBTNsByName As New Dictionary      'dicion�rio pra guardar os bot�es de Reset do form, indexando pelo nome do controle
     '----------------
      
     '----------------
     ' dict pra cada Form pra guardar o objeto de classe com par�metros dos Controles
     '  que v�o alternar entre Enable/Disable
       Public dictCtrlEnblDsblParams As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os par�metros
     
     ' objeto de classe pra armazenar par�metros dos Controles
       Public clObjCtrlsEnblDsblParams As cls_07aCtrlsEnblDsblParams   'declara��o do objeto de classe
             
     ' -->  dicion�rio, dentro do objeto de classe pra armazenar os par�metros separados por Enabled e Disabled
            Public dictParamByLckdStatus As New Dictionary

     ' --->    objeto de classe pra armazenar par�metros dos Controles
               Public clObjLckdStatusParam As cls_07bLckdStatusParams   'declara��o do objeto de classe
     '----------------
       
     '----------------
     ' objeto de classe para armazenar par�metros utilizados nas rotinas de abertura de formul�rio
       Public clObjFormOpenParams As cls_09cParamsToOpenForms
       
     '----------------
     ' dict pra guardar par�metros de comportamento de Triggers e outros controles, por form
       Public dictCtrlBehvrParams As New Dictionary
     
     ' objeto de classe pra armazenar par�metros dos Listbox dos formul�rios
       Public clObjCtrlBehvrParams As cls_11aCtrlBehvrParams   'declara��o do objeto de classe
     '----------------

     '----------------
     ' dict pra guardar par�metros de Formul�rios
       Public dictFormsParams As New Dictionary
            'key   > Grupo de Filtragem da Listbox
            'valor > Objeto de Classe clObjTargtCtrlParam com todos os par�metros
     
     ' objeto de classe pra armazenar par�metros dos formul�rios
       Public clObjFormsParams As cls_09aFormsParams   'declara��o do objeto de classe
     
' ----> objeto de classe interno � Classe [ clObjFormsParams ] pra armazenar alertas de erro de carga de formul�rios
        Public clObjStatusBarWarn As cls_09bFormsParamsLoadLogTxt   'declara��o do objeto de classe
     '----------------


     '----------------
     ' dict pra carregar Controles cujos eventos devem ser monitorados
       Public dictCtrlsEvents As New Dictionary
     
     ' objeto de classe pra armazenar par�metros dos Listbox dos formul�rios
       Public clObjCtrlsEvents As cls_10aCtrls_Events   'declara��o do objeto de classe
     '----------------


     '----------------
     ' dict pra guardar os DataFields dos controles, por form
       Public dictFormDataFlds00Ctrls As New Dictionary   'dicion�rio com os controles tipo DataField do [ form ]
       Public dictFormDataFlds01Grps As New Dictionary   'dicion�rio com os Grupos associados aos controles tipo DataField
     
     
     ' objeto de classe pra armazenar par�metros dos Listbox dos formul�rios
       Public clObjCtrlDataFieds As cls_04aCtrlsDataFields
     '----------------

      
     '----------------
      'dict para guardar as consultas originais dos controles [ Listbox ] e [ Combobox ]
      Public dictFormQrysCtrls As New Dictionary
     '----------------


     '----------------
     ' dict pra guardar os DataFields dos controles, por form
       Public dictFormCommButtons As New Dictionary
     
     ' objeto de classe pra armazenar par�metros dos Listbox dos formul�rios
       Public clObjCommButtons As cls_12aCommButtonParams
     '----------------




'----------------------------------------------
'--  de Cor ----
'    'Types: declara��o de vari�veis para a fun��o LongDecToRGB_HEX_HSL
      Type MyColors
         CorAccessGBlng As Long
         CorRGBgbStr As String
         CorHEXgbStr As String
         CorHSLgbStr As String
      End Type
      Public CoRes As MyColors
    
'-   'Types: declara��o de vari�veis para a fun��o que identifica o RGB da cor from DEC
      Type RGBcolor
         Red As Long
         Green As Long
         Blue As Long
      End Type
      Public GetRGB As RGBcolor

'-   'Types: declara��o de vari�veis para a fun��o que identifica o valor HSL
      Type HSLcolor
         lHue As Long
         lSat As Long
         lLum As Long
         sTintButton As String
      End Type
      Public GetHSL As HSLcolor


'----------------------------------------------
'--  de ... ----
     'Types: declara��o de vari�veis pra armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
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
     'Types: declara��o de vari�veis pra armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
      Type vCtrlPrmissGrnted
          bPermissionGrated As Boolean
          sCtrlNewTipText As String
          
      End Type
      Public GetPrmissGrntedType As vCtrlPrmissGrnted
'-----------------------------------
'----------------------------------------------
      
'--  de ... ----
     'Types: declara��o de vari�veis pra
      Type vCheckQryFld
          bFoundQryFld As Boolean     'Verdadeiro se o campo ora analisado estiver presente no Grid da consulta
          sQry As String              'usado apenas pra recuperar o nome consulta ora analisada, como String pra ser exibido na msg de erro
          
      End Type
      Public NstdVarQryFld As vCheckQryFld

'-----------------------------------
'----------------------------------------------
      
''--  de ... ----
'     'Types: declara��o de vari�veis pra
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
'     'Types: declara��o de vari�veis para armazenar o SQL de cada controle e o nome do controle que deve exibir a contagem
'     ' User Defined Types
'      Type vUserType
'          sCtrlWhere As String            'string SQL WHERE com o texto que deve ser usado pra pesquisa dos dados
'          sCtrlReCntCption As String     'string para exibi��o na contagem de registros
'      End Type
'      Public NstdVar As vUserType
'
'
'       vUserType
'       NstdVar             NstdVar = GetFunction(var01 as ...
'       GetFunction         Public Function GetSQL_01_Chkbox(iFrmIndexID As Integer, ...


' ---------------------------------------------------------------
'      'chamada da fun��o, na rotina principal, pra obten��o dos valores
'      '  NstdVar = GetFunction(iFrmIndexID, cTriggCtrl, ...)
'      '
' ---------------------------------------------------------------
'      ' fun��o pra obten��o dos valores
'      '   Public Function GetFunction(iFrmIndexID As Integer, cTriggCtrl As Control, ...) As vUserType
'      '     GetFunction.sCtrlWhere = sWhere
'      '     GetFunction.sCtrlReCntCption = sReCntCptn
' ---------------------------------------------------------------

' ---------------------------------------------------------------
'      '
'      'retorno � chamada, usando os valores
'      '   vA = NstdVar.sCtrlWhere
'      '   vB = NstdVar.sCtrlReCntCption
' --  Modelo ----
' ---------------------------------------------------------------


                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
''
'' A verificar se deve ser levado de volta pra �rea de Declara��o de vari�veis

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
