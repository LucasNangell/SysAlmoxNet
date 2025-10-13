Attribute VB_Name = "M�dulo 04a - FormsDictsBuild"
Option Compare Database
Option Explicit


Public Sub pbSub00_UserPermissionsDictBuild(vLoginStR As Variant)
    'Armazena os dados e as permiss�es do usu�rio cujo Login corresponda ao vLoginStR
    Dim vA, vB, vC, vD
    'Dim dDicT As Dictionary
    Dim vKey As Variant
    Dim cCtrL As Control
    Dim sWhere As String
    Dim iInT As Integer
    Dim rsTbE As Recordset
    Dim iUserPermission As Integer
    Dim sStR1 As String, sStR2 As String
    Dim sQuerY As String
    Dim sUserLoginSrch As String
    Dim iCounT As Integer
    Dim lngFoundRecs As Long
    Dim sLoadLogWarn As String
'Stop

    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

    'Cria o objeto [ clObjUserParams ] da Classe [ cls_08aLoggedUserParams ] pra guardar os dados do Usu�rio
    Set clObjUserParams = New cls_08aLoggedUserParams
    
'Stop
    '------------------------------------------------
    'Inclui valores de inicializa��o pro caso dos dados reais n�o serem identificados pela Rotina
    clObjUserParams.lngUserID = 0
    clObjUserParams.sUserLogin = "0000"
    clObjUserParams.sUserName = "Ignorado"
    clObjUserParams.iUserSetor = 0
    clObjUserParams.sUserSetor = "Ignorado"
    Set clObjUserParams.dictUserPermissions = New Dictionary
    clObjUserParams.dictUserPermissions.Add "99", "N�vel n�o identificado"


'Stop
    '-----------------------------------------------------------------------------------
    '------------------------------------------------
    'Tenta recuperar os dados efetivos do usu�rio logado
    ' se alguma coisa der errado permanece com os valores de inicializa��o padr�o j� armazenados
    ' ap�s exibi��o de mensagem de erro apropriada
    
    '------------------------------------------------
    'Define par�metros pra consultar os [ dados de Usu�rio ]
    sQuerY = "qry_00(00)aSysUsers"
    sUserLoginSrch = "UserLoginStR"
    
    'Trata a String com o Login de usu�rio
    vA = Replace(vLoginStR, "P_", "")
    vLoginStR = vA
    
    '------------------------------------------------
    'Monta a condi��o Where a ser usada na Consulta pra recuperar os dados do usu�rio
    iInT = IIf(TypeName(vLoginStR) = "Integer", 1, IIf(TypeName(vLoginStR) = "String", 2, 0))
    'vA = TypeName(vLoginStR)
'Stop
    
    If iInT = 0 Then
        'O valor de vLoginStR n�o � nem um Integer nem uma String
        MsgBox "erro n�o previsto"
'Stop
    
    Else
        
        vA = IIf(iInT = 1, " = " & vLoginStR, " LIKE " & """" & vLoginStR & """")
        
    End If
    
    sWhere = "[" & sUserLoginSrch & "]" & vA
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Get User"
'Stop
    '------------------------------------------------
    'Verifica se o usu�rio logado foi identificado no sistema
    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    rsTbE.Filter = sWhere
    Set rsTbE = rsTbE.OpenRecordset
    
    lngFoundRecs = rsTbE.RecordCount
    If lngFoundRecs > 0 Then
        rsTbE.MoveLast
        lngFoundRecs = rsTbE.RecordCount
    
    Else
        sStR1 = "O usu�rio [ " & vLoginStR & " ] n�o foi localizado" & vbCr & "na Tabela [ " & sQuerY & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = vbCr & " Funcionalidades do sistema que exijam permiss�o" & vbCr & " de usu�rio diferenciada n�o estar�o acess�veis."
        
        Call msgboxErrorAlert(sStR1, sStR2)
        
        Exit Sub
    
    End If
    
    'Confirma que n�o haja mais de um registro referente ao usu�rio
    If lngFoundRecs > 1 Then
        sStR1 = "O usu�rio [ " & vLoginStR & " ] foi localizado em duplicidade" & vbCr & "na Tabela [ " & sQuerY & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = vbCr & " Funcionalidades do sistema que exijam permiss�o" & vbCr & " de usu�rio diferenciada n�o estar�o acess�veis."
        
        Call msgboxErrorAlert(sStR1, sStR2)
        
        Exit Sub
    
    End If

    '------------------------------------------------
    'Substitui os dados do usu�rio inicialmente armazenados pelos dados reais identificados na tabela de Usu�rios
    ' Se houver erro significa que o campo de dados n�o foi localizado na consulta [ sQuerY ]

    'se EOF ou BOF forem TRUE signfica que a String procurada
    ' em  [ rsTbE.Filter ] n�o foi localizada na Consulta
    If Not (rsTbE.EOF And rsTbE.BOF) Then
        rsTbE.MoveFirst
        Do Until rsTbE.EOF = True
'Stop
            'Substitui no objeto [ clObjUserParams ] da classe [ cls_08aLoggedUserParams ] os dados do usu�rio
            clObjUserParams.lngUserID = rsTbE.Fields("UserID")
            clObjUserParams.sUserLogin = rsTbE.Fields("UserLoginStr")
            clObjUserParams.sUserName = rsTbE.Fields("UserName")
            clObjUserParams.iUserSetor = rsTbE.Fields("SetorIDfk")
            clObjUserParams.sUserSetor = rsTbE.Fields("SetorDescri�ao")
            
            'Move to the next record
            rsTbE.MoveNext
        
        Loop
    
    End If
    
    rsTbE.Close 'Close the recordset
    Set rsTbE = Nothing 'Clean up
    
'Stop

'If gBbDepurandoLv01b Then MsgBox "teste - consulta das permiss�es"
'Stop
    '-----------------------------------------------------------------------------------
    '------------------------------------------------
    'Tenta recuperar as permiss�es efetivas do usu�rio logado

    '------------------------------------------------
    'Define par�metros pra consultar as [ permiss�es do Usu�rio ]
    sQuerY = "qry_10(00)cSysUsersPrmissSetorJct(Edt)"
    sUserLoginSrch = "UserLoginStR"
    
    '------------------------------------------------
    'Monta a condi��o Where a ser usada na Consulta pra recuperar os dados do usu�rio
    iInT = IIf(TypeName(vLoginStR) = "Integer", 1, IIf(TypeName(vLoginStR) = "String", 2, 0))
    'vA = TypeName(vLoginStR)
'Stop
    
    If iInT = 0 Then
        'O valor de vLoginStR n�o � nem um Integer nem uma String
        MsgBox "erro n�o previsto"
Stop
    
    Else
'vLoginStR = "6221"
        vA = IIf(iInT = 1, " = " & vLoginStR, " LIKE " & """" & vLoginStR & """")
        
    End If
    
    sWhere = "[" & sUserLoginSrch & "]" & vA
    
'Stop
    '------------------------------------------------
    'Verifica se o usu�rio logado foi identificado na tabela de permiss�es
    ' se ele n�o tiver permiss�es ser� mantido o valor ZERO - N�o identificado
    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    rsTbE.Filter = sWhere
    Set rsTbE = rsTbE.OpenRecordset
    
    
    '------------------------------------------------
    'Substitui as permiss�es do usu�rio inicialmente armazenados pelas permiss�es efetivas identificados na tabela de Permiss�es
    ' Se houver erro significa que o campo de dados n�o foi localizado na consulta [ sQuerY ]
    
    'se EOF ou BOF forem TRUE signfica que a String procurada
    ' em rsTbE.Filter n�o foi localizada na Consulta

'Stop

    'Se o Usu�rio n�o tiver permiss�es n�o passa pelo loop e mant�m apenas a Permiss�o "99" inclu�da no in�cio da rotina
    'clObjUserParams.dictUserPermissions.RemoveAll
    If Not (rsTbE.EOF And rsTbE.BOF) Then
        rsTbE.MoveFirst
        Do Until rsTbE.EOF = True
'Stop
            
            'Substitui no Dict [ dictUserPermissions ] do objeto [ clObjUserParams ] da classe [ cls_08aLoggedUserParams ]
            ' as permiss�es do usu�rio
            ' se houver erro significa que o campo de dados n�o foi localizado na consulta [ sQuerY ]
            vA = rsTbE.Fields("UserLoginLevlsIDfk")
            vB = rsTbE.Fields("UserLoginLevelDescri�")
            
            
            
            If clObjUserParams.dictUserPermissions.Exists(vA) Then
                'set
Stop
            Else
            
                clObjUserParams.dictUserPermissions.Add vA, vB
            
            End If
            
            'Move to the next record
            rsTbE.MoveNext
        
        Loop
        
        'Como foram encontradas e armazenadas permiss�es efetivas pro Usu�rio
        ' apaga a permiss�o de inicializa��o
        clObjUserParams.dictUserPermissions.Remove "99"
        
    End If
    rsTbE.Close 'Close the recordset
    Set rsTbE = Nothing 'Clean up


Erro_FrM_SaiR:
    Exit Sub

FrM_ErrorHandler:
'    Stop
    If (Err.Number = 3265) Then    'campo de dados n�o localizado
        vA = "UserPermissionID"
        sStR1 = "Consulta/Tabela:  [ " & sQuerY & " ]" & vbCr & "Campo de Tabela: " & " [ " & vA & " ]" & vbCr & "-------------------------------------------------------------------------------"
        vB = "O Campo da Consulta/Tabela n�o foi localizado devido a erro" & vbCr & " na Rotina e o usu�rio logado n�o pode ser identificado."
        sStR2 = vB & vbCr & vbCr & " Funcionalidades do sistema que exijam permiss�o" & vbCr & " de usu�rio diferenciada n�o estar�o acess�veis."
        vC = " Erro [ " & Err.Number & " ] "

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vC)

'Stop

    ElseIf (Err.Number = 3078) Then     'tabela n�o localizada
        sStR1 = "Consulta:  [ " & sQuerY & " ]" & vbCr & "-------------------------------------------------------------------------------"
        vB = "A Consulta/Tabela n�o foi localizada no sistema e o" & vbCr & " usu�rio logado n�o poder� ser identificado."
        sStR2 = vB & vbCr & vbCr & " Funcionalidades do sistema que exijam permiss�o" & vbCr & " de usu�rio diferenciada n�o estar�o acess�veis."
        vC = " Erro [ " & Err.Number & " ] "

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vC)
Stop
        Resume Erro_FrM_SaiR

    Else
        'Exibe o c�digo de erro
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
        Application.Quit

    End If


End Sub



Public Sub pbSub10_EventsDictBuild(sForM, sEventCtrL)
    Dim vA, vB
    
'Stop
    If Not IsObject(dictCtrlsEvents(sForM)) Then Set dictCtrlsEvents(sForM) = New Dictionary

    'Se o controle ainda n�o foi inclu�do no dicion�rio
    If Not dictCtrlsEvents(sForM).Exists(sEventCtrL) = True Then
       
       'Cria um novo objeto [ clObjCtrlsEvents ] da Classe [ cls_10aCtrls_Events ] pra ser inclu�do no [ dictCtrlsEvents(sForM) ]
        Set clObjCtrlsEvents = New cls_10aCtrls_Events
        dictCtrlsEvents(sForM).Add sEventCtrL, clObjCtrlsEvents
        
        clObjCtrlsEvents.sCtrlName = sEventCtrL
        'A inicializa��o dos controles ser� feita posteriormente, na abertura do formul�rio
    
    End If

End Sub


Public Function pbSub20_TargtCtrlsDictStartUp(fForM As Form) As Boolean

    Dim vA, vB, vC
    Dim sForM As String
    Dim cTrgtCtrl As Control
    
    Dim sTrgtCtrl As String
    Dim sLstbxTag As String
    Dim vSplittedTAG As Variant
    Dim iTagSection As Integer
    Dim dDicT As Dictionary
    Dim sStR1 As String, sStR2 As String

    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

'Stop
    'Passa por todos os objetos e ao encontrar Listbox e Combobox chama
    ' rotina pra guardar as propriedades dos controles
    ' no Dict [ dictListboxParams(sForM) ] seus par�metros e propriedades

    sForM = fForM.Name

'Stop
    'Loop pra localizar os Listboxes do [ Form ] e incluir nos Diversos Dicts
    For Each cTrgtCtrl In fForM.Controls
        
        sTrgtCtrl = cTrgtCtrl.Name
        sLstbxTag = cTrgtCtrl.Tag
        
        Select Case cTrgtCtrl.ControlType
        
            Case acListBox, acComboBox  'Avalia apenas esses dois tipos de controle pois s�o os que podem ser TargtCtrls
'Stop
                'Se o Controle for um [ TargtCtrl ] chama rotina pra guardar os par�metros no Dict.
                ' Devem ser armazenados diversos par�metros necess�rios pra alterar o SQL do controle
                ' e fazer a filtragem
                If InStr(1, sLstbxTag, "TrgtCtrl") > 0 Then
'Stop
                    
vA = "----- pbSub20_TargtCtrlsDictStartUp --------------------------------------------" & vbCr & vbCr & "Avalia se [ " & sTrgtCtrl & " ] tem a TAG necess�ria pra "
vB = vbCr & "inclus�o no dict [ dictFormFilterGrpsTrgts(sForm) ]" & vbCr & vbCr & "TAG [" & Chr(160) & sLstbxTag & Chr(160) & "]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop

                'Chama rotina pra montar o dicion�rio de [ TrgtCtrls ]
                '-------------------------------------------------------------------------------------------------------------
                '----------------------------------------------
                    'Separa os par�metros do controle em quatro se��es
                    vSplittedTAG = Split(sLstbxTag, "-")
                    iTagSection = 1
                    
                    'Avalia a 2a se��o com par�metros de TrgtCtrl
                    If vSplittedTAG(iTagSection - 1) <> "" Then
                        
'pbSub20_TargtCtrlsDictStartUp
vA = "----- pbSub20_TargtCtrlsDictStartUp --------------------------------------------" & vbCr & vbCr & "Chama " & "[" & Chr(160) & "pbSub21_TargtCtrlsDictBuild" & Chr(160) & "] pra inclus�o de"
vB = vbCr & "[ " & sTrgtCtrl & "  ] no dict [ dictFormFilterGrpsTrgts(sForm) ]" & vbCr & " " & vbCr & " "
'MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop
                        
'Stop
                        If gBbDebugOn Then Debug.Print "  " & cTrgtCtrl.Name
                        On Error GoTo -1
                        Call pbSub21_TargtCtrlsDictBuild(vSplittedTAG(iTagSection - 1), cTrgtCtrl)
                        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
'Stop


If gBbDepurandoLv01b Then MsgBox "----- pbSub20_TargtCtrlsDictStartUp --------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub21_TargtCtrlsDictBuild ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
                    Else
If gBbDepurandoLv01b Then MsgBox "----- pbSub20_TargtCtrlsDictStartUp --------------------------------------------" & vbCr & vbCr & "[ " & sTrgtCtrl & "  ] desconsiderado como Target"
If gBbDepurandoLv01b Then Stop
                    
                    End If
                '-------------------------------------------------------------------------------------------------------------
                '----------------------------------------------
                
                End If
                    
        End Select
    
    Next cTrgtCtrl
'Stop

    'Se n�o tiver encontrado [ TargtCtrls ] define como [ True ] pra que mais adiante
    ' n�o sejam carregados os dicion�rios dos [ TriggCtrls ]
    If IsObject(dictFormFilterGrpsTrgts(sForM)) Then pbSub20_TargtCtrlsDictStartUp = True
    

'Stop
FrM_Error_SaiR:
    On Error GoTo -1
    Exit Function

FrM_ErrorHandler:
Stop
'    If Err.Number = 9 Then
'        'Matriz n�o cont�m os itens esperados
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = "A [ " & iTagSection & "a se��o ] " & " de par�metros do TriggerCtrl" & vbCr & " n�o foi localizada." & vbCr & vbCr & " N�o ser� poss�vel filtrar esse TargetCtrl."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'        'Inclui o erro no dict de Logs de Carga do sistema
'        sLoadLogWarn = "A [ " & iTagSection & "a se��o ] " & " de par�metros do TriggerCtrl n�o foi localizada." & vbCrLf & " N�o ser� poss�vel filtrar os seguintes TargetCtrls."
'        Call FormStatusBar01_Bld(sForm, "3rdTagSectionNotFound", sLoadLogWarn, sTrgtCtrL)
'
'        Exit Function
'
'    Else
        'Erro n�o previsto
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
        Resume Next
        Application.Quit

'    End If
End Function

Public Sub pbSub21_TargtCtrlsDictBuild(vTagSection As Variant, cListBox As Control)

    Dim vA, vB, vC, vD
    Dim vTagSectionParams As Variant
    
    Dim iWhere As Integer
    Dim bBoL As Boolean
    Dim sParam As String
    Dim iFilGrp As Integer
    Dim sFilGrp As String
    Dim sTrgtCtrl As String
    Dim sForM As String
    Dim dDicT As Dictionary
    Dim sRecCntCtrl As String
    Dim cRecCnt As Control
    Dim sStR1 As String, sStR2 As String
    Dim sFilGrpTag As String
    Dim sLoadLogWarn As String
    
    sTrgtCtrl = cListBox.Name
    sForM = cListBox.Parent.Name
    

'MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Avalia [ " & sTrgtCtrl & " ] pra inclus�o no [ TargtDict ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
'Stop
    
    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
    
'Stop '----------------------------
    'sTempStR = sTagSection
    
    'sTagSection = "teste"
    'sTagSection = sTempStR
    vTagSectionParams = Split(vTagSection, ".")    'vSplittedTag(1)
    
    '-------------------------------------------------------------------------------------------------------
    'Recupera os par�metros do controle informados na TAG
    '-------------------------------------------------------------------------------------------------------
'Stop
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    sParam = "Grp"
        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "O TargetCtrl a seguir n�o est� associado a nenhum grupo  de filtragem e n�o poder� ser pesquisado."
        
        'Mensagem de erro a ser exibida em tela
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = " O [ grupo de filtragem ] do TargetCtrl n�o foi informado" & vbCr & "  e ele n�o poder� ser pesquisado."
        
        On Error GoTo -1
        sFilGrp = GetTagParams(sParam, vTagSectionParams, , True, "", 1, , True, sStR1, sStR2, True, "MissingTrgtFilGrp", cListBox, sLoadLogWarn)

'MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Filtergrp: [ " & sFilGrp & " ]" & vbCr & " " & vbCr & " "
'Stop

        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        If sFilGrp = "" Then Exit Sub
'Stop
    
    'O Listbox est� associado a um Grupo de Filtragem
    ' ent�o inicia a montagem do dicion�rio [ dictFormFilterGrpsTrgts ] com par�metros do TrgtCtrl
    
    'Verifica se o Listbox tem um [ cRecCnt ] associado
    sParam = "RCnt>"
        'Mensagem de erro a ser exibida em tela e registrada no Log de carga do sistema
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Listbox: " & "       [ " & sTrgtCtrl & " ]" & vbCr & "Contr. de Contag. de Regs: " & " [ " & sRecCntCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = " O controle de contagem de registros indicado como" & vbCr & "  contador de registros do Listbox n�o foi localizado." & vbCr & "  Sua contagem de registros n�o ser� exibida."
        
        
        sLoadLogWarn = "O controle de contagem de registros indicado como contador de registros do Listbox [ " & sTrgtCtrl & " ] n�o foi localizado." & vbCrLf & "Sua contagem de registros n�o ser� exibida"

        On Error GoTo -1
        sRecCntCtrl = GetTagParams(sParam, vTagSectionParams, cListBox, False, "", , , True, sStR1, sStR2, True, "RCntNotFound", cListBox, sLoadLogWarn)
        'sRecCntCtrl = GetTagParams(sParam, vTagSectionParams, , "", , , sStR1, sStR2)
'Stop
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
    On Error GoTo -1
    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------
    'Cria uma nova varia��o do dicion�rio pro Formul�rio corrente, caso ele ainda n�o tenha sido criado
    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------

'MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Filtergrp: [ " & sFilGrp & " ]" & vbCr & " " & vbCr & " "
'Stop
    
    If Not IsObject(dictFormFilterGrpsTrgts(sForM)) Then Set dictFormFilterGrpsTrgts(sForM) = New Dictionary
    'Set dDicT = dictFormFilterGrpsTrgts(sForM)


'MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Avaliando o TrgtCtrl [ " & sTrgtCtrl & " ]" & vbCr & "Inclui o Grupo [ " & sFilGrp & " ] em [ dictFormFilGrpsEnDsAllCtrls(sForm) ]" & vbCr & " " & vbCr & " "
'If gBbDepurandoLv01c Then Stop
'Stop


    'Trecho que prepara para execu��o posterior do [ Enbl/Dsbl ]
    If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM) = New Dictionary
    If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp) = New Dictionary
    If Not dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Exists(cListBox.Name) Then dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Add cListBox.Name, sFilGrp
    
'Stop
    'Verifica se j� foi criado um dict do [ sFilGrp ] ora avaliado
    If Not IsObject(dictFormFilterGrpsTrgts(sForM)) Then Set dictFormFilterGrpsTrgts(sForM) = New Dictionary
    If Not IsObject(dictFormFilterGrpsTrgts(sForM)(sFilGrp)) Then Set dictFormFilterGrpsTrgts(sForM)(sFilGrp) = New Dictionary
    
''------------------------------------------------------------------------------------------------------
''parei aqui: checar por que o item sFilGrp � inclu�do no dictFormFilterGrpsTrgts quando passa pelo If
'    If dictFormFilterGrpsTrgts.Exists(sForM) And dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp) Then
'
'Stop
'    End If
'
'
'    For Each vA In dictFormFilterGrpsTrgts(sForM)
'
'Stop
'    Next vA
    
'Stop
    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------
    'Inicia a cria��o do dictFormFilterGrpsTrgts(sForM) e simultaneamente a cria��o do
    ' "sub dicion�rio"  [ dictFormFilterGrpsTrgts(sForM)(sFilGrp) ] pra inclus�o
    ' dos [ TrgtCtrls ] associados ao Grupo
    If dictFormFilterGrpsTrgts(sForM)(sFilGrp).Exists(sTrgtCtrl) = True Then
        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)(sTrgtCtrl)
    
    Else
        'Cria um novo objeto [ clObjTargtCtrlParam ] da Classe [ cls_01aTargtCtrlParams_Evnts ] pra ser inclu�do no Dict
        Set clObjTargtCtrlParam = New cls_01aTargtCtrlParams_Evnts

        'Adiciona um novo item no dicion�rio [ dictFormFilterGrpsTrgts(sForM)(sFilGrp) ] e guarda nele o objeto [ clObjTargtCtrlParam ]
        ' com os respectivos par�metros do [ TrgtCtrl ] definidos na classe [ cls_01aTargtCtrlParams_Evnts ]
        dictFormFilterGrpsTrgts(sForM)(sFilGrp).Add sTrgtCtrl, clObjTargtCtrlParam
        
        If Not IsObject(dictTrgtCtrlsFilterGrps(sForM)) Then Set dictTrgtCtrlsFilterGrps(sForM) = New Dictionary
        dictTrgtCtrlsFilterGrps(sForM).Add cListBox.Name, sFilGrp
        
        'Constru��o do dicion�rio [ clObjTargtCtrlParam.dictQryFields ] com os campos da consulta do [ TrgtCtrl ]
        Call PbSbBuildDictFieldsInQryTrgtCtrl(sForM, cListBox, sFilGrp)
    
    End If
'Stop

    'Chama a rotina pra recuperar o SQL do [ cListbox ]
    On Error GoTo -1
    sGbQrySQLstr = pbSub22_GetTargtCtrlsSQL(cListBox)
'Stop
    With clObjTargtCtrlParam
        'Atribui ao Listbox os par�metros do Listbox esperados pela Classe [ cls_01aTargtCtrlParams_Evnts ]
        .sTargtCtrlName = cListBox.Name
        .sRecCntCtrlName = sRecCntCtrl
        .sFilGrp = sFilGrp
        
        'dictFormFilterGrpTrgts
        
        
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Get SQL"
'Stop

'pbSub21_TargtCtrlsDictBuild
If gBbDepurandoLv01c Then MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub22_GetTargtCtrlsSQL ] e salva o SQL" & vbCr & "de [ " & sTrgtCtrl & " ] em [ dictFormFilterGrpsTrgts(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
        On Error GoTo -1
        .sClsLstbxSQL_aSELECT = sGbQrySQLstr.sLstbxSQL_aSELECT
        .sClsLstbxSQL_bFROM = sGbQrySQLstr.sLstbxSQL_bFROM
        .sClsLstbxSQL_cWHERE = sGbQrySQLstr.sLstbxSQL_cWHERE
        .sClsLstbxSQL_dOrderBy = sGbQrySQLstr.sLstbxSQL_dOrderBy
        .sClsLstbxSQL_eMAIN = sGbQrySQLstr.sLstbxSQL_eMAIN
    
    End With
    
'    'ANTIGO
'    '-------------------------------------------------------------------------------------------------------
'    '-------------------------------------------------------------------------------------------------------
'    'Se o Grupo de Filtragem j� tiver sido inclu�do no [ Dict ] significa que
'    ' mais de um [ TargtCtrl ] est� associado a esse Grupo de Filtragem
'    ' nesse caso desconsidera esse �ltimo [ TargtCtrl ]
'    If dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp) = True Then
''Stop
'        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)
'
'        vA = clObjTargtCtrlParam.sTargtCtrlName
'        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Grupo de Filtragem: " & " [ " & sFilGrp & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = "Listbox [ " & sTrgtCtrl & " ]" & " associado como TargtCtrl do Grupo," & vbCr & " em duplicidade com o Listbox [ " & vA & " ]." & vbCr & vbCr & " O Listbox [ " & sTrgtCtrl & " ] ser� desconsiderado" & vbCr & " e n�o poder� ser filtrado."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'
'        'Inclui o erro no dict de Logs de Carga do sistema
'        sLoadLogWarn = "Os seguintes Listboxes foram associados como TargtCtrl do Grupo [ " & sFilGrp & " ] em duplicidade com o Listbox [ " & vA & " ]." & vbCrLf & " Esses Listboxes ser�o desconsiderados e n�o poder�o ser filtrados."
'        Call FormStatusBar01_Bld(sForM, "DupTargtsIn" & "_" & sFilGrp, sLoadLogWarn, sTrgtCtrl)
'
'Stop
'    Else
'        'Cria um novo objeto [ clObjTargtCtrlParam ] da Classe [ cls_01aTargtCtrlParams_Evnts ] pra ser inclu�do no Dict
'        Set clObjTargtCtrlParam = New cls_01aTargtCtrlParams_Evnts
'
'        'Adiciona um novo item no dicion�rio [ dictFormFilterGrpsTrgts ] e guarda nele o objeto [ clObjTargtCtrlParam ]
'        ' com os respectivos par�metros do Grupo de Filtragem definidos na classe [ cls_01aTargtCtrlParams_Evnts ]
'        dictFormFilterGrpsTrgts(sForM).Add sFilGrp, clObjTargtCtrlParam
'
'        If Not IsObject(dictTrgtCtrlsFilterGrps(sForM)) Then Set dictTrgtCtrlsFilterGrps(sForM) = New Dictionary
'        dictTrgtCtrlsFilterGrps(sForM).Add cListBox.Name, sFilGrp
'
''MsgBox "build"
''Stop
'        With clObjTargtCtrlParam
'            'Atribui ao Listbox os par�metros do Listbox esperados pela Classe [ cls_01aTargtCtrlParams_Evnts ]
'            .sTargtCtrlName = cListBox.Name
'            .sRecCntCtrlName = sRecCntCtrl
'            .sFilGrp = sFilGrp
'
''MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Get SQL"
''Stop
'            'Chama a rotina pra recuperar o SQL do [ cListbox ]
'            On Error GoTo -1
'            sGbQrySQLstr = pbSub22_GetTargtCtrlsSQL(cListBox)
''Stop
''pbSub21_TargtCtrlsDictBuild
'If gBbDepurandoLv01c Then MsgBox "----- pbSub21_TargtCtrlsDictBuild ---------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub22_GetTargtCtrlsSQL ] e salva o SQL" & vbCr & "de [ " & sTrgtCtrl & " ] em [ dictFormFilterGrpsTrgts(sForM) ]" & vbCr & " " & vbCr & " "
'If gBbDepurandoLv01c Then Stop
'            On Error GoTo -1
'            .sClsLstbxSQL_aSELECT = sGbQrySQLstr.sLstbxSQL_aSELECT
'            .sClsLstbxSQL_bFROM = sGbQrySQLstr.sLstbxSQL_bFROM
'            .sClsLstbxSQL_cWHERE = sGbQrySQLstr.sLstbxSQL_cWHERE
'            .sClsLstbxSQL_dOrderBy = sGbQrySQLstr.sLstbxSQL_dOrderBy
'            .sClsLstbxSQL_eMAIN = sGbQrySQLstr.sLstbxSQL_eMAIN
'
'        End With
'
'    End If
'    '-------------------------------------------------------------------------------------------------------
'    '-------------------------------------------------------------------------------------------------------


'Stop

        ''Teste de acesso aos par�metros armazenados
        '' .atribui a [ dictFormFilterGrpsTrgts(sForM) ] o dicion�rio [ dictListboxParams("frm_01(1)cProdEstoque") ]
        'Set dictFormFilterGrpsTrgts(sForM) = dictListboxParams(fForm.Name)
        '
        '' .verificar se o item j� existe no dicion�rio
        'bBol = dictListboxParams(fForm.Name).Exists(cListbox.Name)
        '
        '' .define o objeto de Classe [ clObjTargtCtrlParam ] como sendo o [ lstTeste2 ] pra exibir seus par�metros
        'Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(cListbox.Name)
        '' .retorna o par�metro
        'vA = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT
'Stop
FrM_Error_SaiR:
    On Error GoTo -1
    Exit Sub

FrM_ErrorHandler:
Stop

'    If Err.Number = 9 Then
'        'Matriz n�o cont�m os itens esperados
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = "O par�mentro " & " [ " & sParam & " ] " & " do TargetCtrl n�o foi localizado." & vbCr & " N�o ser� poss�vel filtrar esse TargetCtrl."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'
'
'        Exit Sub
'
'    ElseIf Err.Number = 2465 Then
'        'Controle do formul�rio n�o foi localizado
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "Listbox: " & "       [ " & sTrgtCtrL & " ]" & vbCr & "Contr. de Contag. de Regs: " & " [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = " O controle de contagem de registros n�o foi localizado." & vbCr & "  N�o ser� poss�vel exibir a contagem de registros associada."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'        sRecCntCtrl = ""
        Resume Next

'    Else
        'Exibe o c�digo de erro
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
        Application.Quit

'    End If

End Sub


Public Function pbSub22_GetTargtCtrlsSQL(cTrgtCtrl As Control, Optional sQuerYname As String) As sLstbxSQLstr
'Public Function fctListboxParam(Optional cCtrl As Control, Optional sQuerY As String, Optional cRecCnt As Control) As cls_01aTargtCtrlParams_Evnts
    Dim vA, vB
    Dim sStR As String
    Dim sFiLnewFrmSELECT  As String
    Dim sFiLnewFrmFROM As String
    Dim sFiLnewFrmWHERE As String
    Dim sFiLnewFrmTmpWHERE As String
    Dim sFiLnewFrmOrderBy As String
    Dim sStartSqL As String
    Dim iWhere As Integer
    Dim qDef As DAO.QueryDef
    Dim vArrTempSQL As Variant
    Dim lgInT As Long
    Dim sTrgtCtrl As String
    Dim sForM As String
    Dim sStR1 As String, sStR2 As String
    
    'Recupera a SQL da Listbox corrente e permite armazenar a SQL no Dicion�rio de Listboxes
    
    'armazena, nas vari�veis globais, o SQL do controle indicado
    ' .sgbFiLnewFrmSELECT
    ' .sgbFiLnewFrmWHERE
    ' .sgbFiLnewFrmOrderBy
    '   isso facilita a mudan�a do WHERE para filtragem do controle
    
    'vA = cTrgtCtrL.Name
    'vB = cTrgtCtrL.Parent.Name
    
    
    sTrgtCtrl = cTrgtCtrl.Name
    sForM = cTrgtCtrl.Parent.Name
    


'pbSub22_GetTargtCtrlsSQL

'MsgBox "----- pbSub22_GetTargtCtrlsSQL -------------------------------------------------" & vbCr & vbCr & "Recupera o SQL de [ " & sTrgtCtrl & " ]" & vbCr & " " & vbCr & " "
'If gBbDepurandoLv01c Then Stop
'Stop
   
    sStR = ""
    sFiLnewFrmSELECT = ""
    sFiLnewFrmWHERE = ""
    sFiLnewFrmTmpWHERE = ""

    sFiLnewFrmOrderBy = ""
    
'Stop
    sStartSqL = cTrgtCtrl.RowSource
    
        
    If sQuerYname <> "" Then
        sStartSqL = sQuerYname
        'sMsgboxLine1 = "A consulta [ " & sQuerYname & " ]   n�o possui uma Consulta ou" & vbCr & " Tabela como fonte de dados"
    
    End If
    'iWhere = InStrRev(sStartSqL, ";")
    iWhere = InStr(sStartSqL, "SELECT")
    
    
'    Debug.Print sStartSQL
'Stop
    '''
'Stop
    If iWhere < 1 Then   'se iWhere < 1 significa que a fonte de dados da lista de resultados n�o � um SQL.
                           'Ent�o deve ser obtido o SQL a partir da fonte de dados do controle
        
        If sStartSqL = "" Then
            
            If gBbDebugOn Then Debug.Print sStartSqL
            
            sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Listbox: " & "       [ " & sTrgtCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
            sStR2 = "A SQL do Listbox est� vazia." & vbCr & "  Este TrgtCtrl n�o exibir� registros" & vbCr & "  e n�o poder� ser filtrado"
            vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
            
            Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
Stop
            Exit Function
        
        Else
        
            If gBbDebugOn Then Debug.Print sStartSqL
            
            
            Set qDef = CurrentDb.QueryDefs(sStartSqL)
            sStartSqL = qDef.sql
            
            'iWhere = InStr(sStartSqL, "SELECT")
            
        End If
        iWhere = InStrRev(sStartSqL, ";")
        
    
    'Se iWhere > 0 significa que  foi encontrada a express�o "SELECT" na fonte de dados do controle
    ' nesse caso deve ser exibido alerta que a fonte de dados do [ TrgtCtrl ] que deveria ser o nome de uma consulta foi substitu�da
    ' por um SQL
'    Else
'
'        sStr1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "Listbox: " & "       [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStr2 = "A fonte de dados do Listbox est� corrompida." & vbCr & "  Este TrgtCtrl n�o exibir� registros" & vbCr & "  e n�o poder� ser filtrado"
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStr1, sStr2, vbExclamation, vA)
'Stop
'        Exit Function
        
    End If




'Stop

    'remove o ; ao final do SQL
    If iWhere > 0 Then sStartSqL = Mid(sStartSqL, 1, iWhere - 1)
    'Debug.Print sStartSqL
    '----------------------------------
'Stop
    'inclui ; como separa��o entre os quatro trechos poss�veis da SQL
    sStartSqL = Replace(sStartSqL, "FROM ", ";FROM ")
    'Debug.Print sStartSqL
    
    sStartSqL = Replace(sStartSqL, "WHERE ", ";WHERE ")
'    Debug.Print sStartSQL
    
    sStartSqL = Replace(sStartSqL, "ORDER BY ", ";ORDER BY ")
    'Debug.Print sStartSqL

'Stop
    'carrega na Matriz, cada um dos trechos localizados
    vArrTempSQL = Split(sStartSqL, ";")
'Stop
       
'Stop  'Ctrl point
    vA = UBound(vArrTempSQL)
    For lgInT = 0 To UBound(vArrTempSQL)
        
        'iWhere = InStrRev(sStartSQL, ";")
        
        sStR = Trim(Left(vArrTempSQL(lgInT), 8))
        
        vA = InStrRev(sStR, "SELECT")
            If vA > 0 Then sFiLnewFrmSELECT = vArrTempSQL(lgInT)
        
        vA = InStrRev(sStR, "FROM")
            If vA > 0 Then sFiLnewFrmFROM = vArrTempSQL(lgInT)
        
        vA = InStrRev(sStR, "WHERE")
            If vA > 0 Then sFiLnewFrmWHERE = vArrTempSQL(lgInT)
        
        vA = InStrRev(sStR, "ORDER")
            If vA > 0 Then sFiLnewFrmOrderBy = vArrTempSQL(lgInT)

    Next lgInT
'Stop
    
    pbSub22_GetTargtCtrlsSQL.sLstbxSQL_aSELECT = sFiLnewFrmSELECT
    pbSub22_GetTargtCtrlsSQL.sLstbxSQL_bFROM = sFiLnewFrmFROM
    pbSub22_GetTargtCtrlsSQL.sLstbxSQL_cWHERE = sFiLnewFrmWHERE
    pbSub22_GetTargtCtrlsSQL.sLstbxSQL_dOrderBy = sFiLnewFrmOrderBy
    pbSub22_GetTargtCtrlsSQL.sLstbxSQL_eMAIN = sFiLnewFrmSELECT & " " & sFiLnewFrmFROM
    
'    sGbFiLnewFrmSELECT = sFiLnewFrmSELECT & " " & sFiLnewFrmFROM
'    sgbFiLnewFrmWHERE = sFiLnewFrmWHERE
'    sgbFiLnewFrmOrderBy = sFiLnewFrmOrderBy
'    Debug.Print sGbFiLnewFrmSELECT
'    Debug.Print sgbFiLnewFrmWHERE
'    Debug.Print sgbFiLnewFrmOrderBy
    
'Stop
End Function


Public Sub pbSub30_TriggCtrlDictStartUp(fForM As Form)

    Dim vA, vB, vC
    'Dim fForM As Form
    Dim sForM As String
    Dim cTriggCtrl As Control
    Dim sTriggCtrl As String
    'Dim sLstbxTag As String
    Dim sCtrlTAG As String
    
    Dim vSplittedTAG As Variant
    Dim iTagSection As Variant
    Dim sStR1 As String, sStR2 As String
    Dim lngEvalFormCtrls As Long, lngEvalTAGedCtrls As Long, lngCtrlsInDict As Long
    Dim bFoundParams As Boolean
    Dim sLoadLogWarn As String
    Dim vKeyFilGrp As Variant, vKeyTrgtCtrls As Variant  'vKeyTrggCtrls As Variant
    Dim iCtrlType As Integer
    Dim sCtrlType As String
    Dim sFilGrp As String
    Dim bCtrlIsTrgg As Boolean
    Dim bCtrlIsTrgt As Boolean
    
'Stop
    'Abre o Form pra recuperar os par�metros dos [ TrggCtrls ] e armazenar
    ' nos Dicts [ dictTrgg... ]

    sForM = fForM.Name

    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

'Stop
    'vA = 1 / 0
    
    'Loop pra localizar os TriggCtrls do [ Form ] e incluir nos Diversos Dicts
    For Each cTriggCtrl In fForM.Controls
        
        sTriggCtrl = cTriggCtrl.Name
'Stop

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Formul�rio [ " & sForM & " ]" & vbCr & "[ " & sTriggCtrl & " ] est� na categoria de [ Triggers ]?"
'If gBbDepurandoLv01c Then Stop
'Stop
               
        Select Case cTriggCtrl.ControlType
        'Avalia os controles pra montagem do dicion�rio de [ Triggers ]
        '-------------------------------------------------------------------------------------------------------------
        '----------------------------------------------
            Case acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox  'Avalia apenas esses tipos de controle pois s�o os que podem ser TriggCtrls
'Stop
                'C�digo para carregar o dicion�rio com as consultas dos controles do tipo [ acListBox ] e [ acComboBox ]
                '---------------------------------------------------------------------------------------------------------
                
                If cTriggCtrl.ControlType = acComboBox Or cTriggCtrl.ControlType = acListBox Then
                    
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & "Carrega a consulta original de" & vbCr
vB = "[ " & sTriggCtrl & " ] no dict [ dictFormQrysCtrls(sForm) ]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
                    
                    If Not IsObject(dictFormQrysCtrls(sForM)) Then Set dictFormQrysCtrls(sForM) = New Dictionary
                    
                    If Not dictFormQrysCtrls(sForM).Exists(cTriggCtrl.Name) Then
                        dictFormQrysCtrls(sForM).Add cTriggCtrl.Name, cTriggCtrl.RowSource
                    
                    End If
                
                End If
                

                'vA = sForM
                sCtrlTAG = cTriggCtrl.Tag
                
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Avalia se [ " & sTriggCtrl & "  ] tem a TAG necess�ria pra inclus�o" & vbCr & "nos dicts [ dictTrgg00GrpsInForm ] e [ dictTrgg01CtrlsInGrp ]"
vB = vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop
                '-------------------------------------------
                'Avalia��o de [ TriggerCtrl ]
                '-------------------------------------------
                
                'Monta o dicion�rio de [ TriggCtrl ] mas antes confirma se existem [ TargtCtrls ] no Form
                If gBbTrgtCtrlsFound Then
                
                If gBbDebugOn Then Debug.Print "Trigger Controls"
                    lngEvalFormCtrls = lngEvalFormCtrls + 1
                    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Dict pbSub06_TriggBuild do controle [ " & sTriggCtrl & "  ]"
'Stop
                    'Se o Controle for mesmo um TriggerCtrl chama rotina pra guardar os par�metros no Dict.
                    ' Devem ser armazenados diversos par�metros necess�rios pra alterar o SQL do controle
                    ' e fazer a filtragem
                    If InStr(1, sCtrlTAG, "Trgg") > 0 Then
                        
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "chama Triggers Dict build, controle [ " & sTriggCtrl & " ]"
'Stop
                        
                    'Chama rotina pra montar o dicion�rio de [ TrggCtrls ]
                    '-------------------------------------------------------------------------------------------------------------
                    '----------------------------------------------
                        'Separa os par�metros do controle em quatro se��es
                        vSplittedTAG = Split(sCtrlTAG, "-")
                        iTagSection = 2
'Stop
                        'Avalia a 1a se��o com par�metros de TrggCtrl
                        If vSplittedTAG(iTagSection - 1) <> "" Then

vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & "Chama [ pbSub31_TriggCtrlDictBuild ] pra inclus�o de" & vbCr
vB = "[ " & sTriggCtrl & " ] nos dicts [ dictTrgg00GrpsInForm ] e" & vbCr & "[ dictTrgg01CtrlsInGrp ]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
                            If gBbDebugOn Then Debug.Print " " & cTriggCtrl.Name
                            On Error GoTo -1
                            
'MsgBox "erro vazio"
'Stop
                            Call pbSub31_TriggCtrlDictBuild(vSplittedTAG(iTagSection - 1), cTriggCtrl)
                            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                            
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "back from Triggers Dict build"
'Stop
                    '----------------------------------------------
                    '-------------------------------------------------------------------------------------------------------------
                        End If
'Stop
    
                    Else
If gBbDepurandoLv01b Then MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "[ " & sTriggCtrl & "  ] N�O tem a TAG pra inclus�o nos dicts" & vbCr & "[ dictTrgg00GrpsInForm ] e [ dictTrgg01CtrlsInGrp ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
                    
                    End If
        '----------------------------------------------
        '-------------------------------------------------------------------------------------------------------------
                End If

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Ctrl: [ " & sTriggCtrl & " ]"
'Stop


vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Avalia se [ " & sTriggCtrl & "  ] tem a TAG necess�ria pra inclus�o" & vbCr & "no dict [ dictFrmResetAreas(sForM) ]"
vB = vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop

                
                sCtrlTAG = cTriggCtrl.Tag
'Stop
                
                '-------------------------------------------
                'Avalia��o de [ ResetArea ]
                '-------------------------------------------
                
                'Verifica se o Controle pertence a alguma [ RstArea ] for mesmo um TriggerCtrl chama rotina pra guardar os par�metros no Dict.
                ' Devem ser armazenados diversos par�metros necess�rios pra alterar o SQL do controle
                ' e fazer a filtragem
                If InStr(1, sCtrlTAG, "RstAr") > 0 Then
'Stop
                    'Separa os par�metros do controle em quatro se��es
                    vSplittedTAG = Split(sCtrlTAG, "-")
                    iTagSection = 4
                
                    'Avalia a 4a se��o com par�metros de [ RstArea ]
                    If vSplittedTAG(iTagSection - 1) <> "" Then
                        
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & "Chama [ pbSub51_RstAreaDictBuild ] pra inclus�o de" & vbCr
vB = "[ " & sTriggCtrl & " ] no dict [ dictFrmResetAreas(sForM)  ]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop
                        If gBbDebugOn Then Debug.Print " " & cTriggCtrl.Name
                        On Error GoTo -1
                        Call pbSub51_RstAreaDictBuild(vSplittedTAG(iTagSection - 1), cTriggCtrl)
                        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
'Stop
                    End If
                
                End If
'Stop

If gBbDepurandoLv01b Then MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "[ " & sTriggCtrl & " ] est� na categoria de [ BehvrCtrls ]?" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
            
            

























'MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Avalia��o de DataFields, controle [ " & sTriggCtrl & " ]"
If gBbDepurandoLv01c Then Stop
'Stop
                '-------------------------------------------
                'Avalia��o de [ DataFields ]
                '-------------------------------------------
                
                'Verifica se o Controle � um [ DataField ]
                If InStr(1, sCtrlTAG, "DataField>") > 0 Then
'Stop
                    'Separa os par�metros do controle em quatro se��es
                    vSplittedTAG = Split(sCtrlTAG, "-")
                    iTagSection = 1
                
                    'Avalia a 4a se��o com par�metros de [ RstArea ]
                    If vSplittedTAG(iTagSection - 1) <> "" Then
                        
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & "Chama [ pbSub51_RstAreaDictBuild ] pra inclus�o de" & vbCr
vB = "[ " & sTriggCtrl & " ] no dict [ dictFrmResetAreas(sForM)  ]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop
                        If gBbDebugOn Then Debug.Print " " & cTriggCtrl.Name
                        On Error GoTo -1
                        Call pbSub71_DataFieldDictBuild(vSplittedTAG(iTagSection - 1), cTriggCtrl)
                        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
'Stop
                    End If

                End If
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
                '-------------------------------------------
                'Avalia��o de [ CtrlsBehvr ]
                '-------------------------------------------
                
                'Chama rotina pra montar o dicion�rio de [ CtrlsBehvr ]
                'Avalia os controles pra montagem do dicion�rio [ dictCtrlBehvrParams(sForM) ]
                '-------------------------------------------------------------------------------------------------------------
                '----------------------------------------------
                iCtrlType = cTriggCtrl.ControlType
                sCtrlType = dictCtrlTypeShort(cTriggCtrl.ControlType)
                
                
                If sCtrlType = "txt" Or sCtrlType = "lst" Or sCtrlType = "cmb" Then     'Avalia apenas esses tipos de controle pra definir comportamento
    
    
    vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Avalia se [ " & sTriggCtrl & "  ] tem a TAG pra inclus�o no dict" & vbCr
    vB = "[ dictCtrlBehvrParams(sForM) ]" & vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
    If gBbDepurandoLv01b Then MsgBox vA & vB
    If gBbDepurandoLv01b Then Stop
    
                    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "checa bFoundParams: [ " & sTriggCtrl & " ]"
'Stop
                    
                    'Antes inicializa os booleans como False
                    bCtrlIsTrgg = False: bCtrlIsTrgt = False
                    
                    'Verifica se [ sTriggCtrl ] tem os requisitos pra ser inclu�do no dict [ dictCtrlBehvrParams(sForM) ]
                    ' se for Listbox ou Combobox e for [ TargetCtrl ] deve ser tamb�m [ TriggCtrl ]
                    If Not cTriggCtrl.Locked Then
                    
                        'Confirma se [ sTriggCtrl ] � [ TriggCtrl ] e tamb�m [ TargetCtrl ]
                        If sCtrlType = "lst" Or sCtrlType = "cmb" Then
                            
'Stop
                            'Confirma se [ sTriggCtrl ] � um [ Trigger ]
                            ' verifica se o dict [dictTrggCtrlsInForm(sForM)] foi criado, o que indica que h� [ TrggCtrls ] carregados
                            vA = IsObject(dictTrggCtrlsInForm(sForM))
                            
                            'Se o dicion�rio de [ TrggCtrls ] n�o existir ou se ele existir mas [ sCtrL ] n�o tiver sido inclu�do, indica que ele N�O � um trigger
                            If vA Then vB = dictTrggCtrlsInForm(sForM).Exists(sTriggCtrl) Else vB = False
                            
                            If vB Then
'Stop
                                'Recupera o [ grupo de filtragem ] do [ TriggCtrl ]
                                If IsObject(dictTrggCtrlsInForm(sForM)(sTriggCtrl)) Then
                                    Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(sTriggCtrl)
                                    sFilGrp = clObjFilGrpsByForm.sFilGrp
                                    bCtrlIsTrgg = True

'parei aqui: checar por que [ bCtrlIsTrgg ] e [ bCtrlIsTrgt ] n�o est�o sendo usados em outros trechos

                                End If
                            
                            End If
                        
                            'Confirma se [ sTriggCtrl ] � um [ TrgtCtrl ]
                            ' confirma se o dict [ dictFormFilterGrpsTrgts(sForM) ] existe, o que indica que h� Grupos de filtragem no [ Form ]
                            If IsObject(dictFormFilterGrpsTrgts(sForM)) Then
                            
                                
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Loop pelos Grupos de Trigg [ " & sTriggCtrl & " ]"
'Stop
On Error GoTo -1
                                'Passa por todos os [ Grupos de Filtragem ] do [ dictFormFilterGrpsTrgts(sForM) ] e verifica se
                                ' o [ sTriggCtrl ] ora avaliado j� foi tamb�m associado ao [ Grupo ] como [ TargetCtrl ]
                                For Each vKeyFilGrp In dictFormFilterGrpsTrgts(sForM)
                                    
                                    For Each vKeyTrgtCtrls In dictFormFilterGrpsTrgts(sForM)(vKeyFilGrp)
                                        
                                        If Not IsEmpty(vKeyTrgtCtrls) And Not vKeyTrgtCtrls = "" Then
                                        
                                            Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(vKeyFilGrp)(vKeyTrgtCtrls)
                                            vA = clObjTargtCtrlParam.sTargtCtrlName
                                            If vA = sTriggCtrl Then bCtrlIsTrgt = True
                                        
                                        End If
                                        
                                    Next vKeyTrgtCtrls
                                    
                                    
                                Next vKeyFilGrp
    
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Trigger � Target"
'Stop
                            End If
                        
                        End If
                    End If
                    
                    '-------------------------------------------------------------------------------------------------------------
                    '----------------------------------------------
                    'Busca a TAG de [ TriggCtrl ]
                    ' se a TAG do controle n�o for vazia verifica se possui a 3a se��o
                    If sCtrlTAG <> "" Then
                        
                        bFoundParams = False
                        'Chama rotina pra montagem do dicion�rio [ CtrlsBehvrParams ]
                        '-------------------------------------------------------------
                        
                        'Separa os par�metros do controle em quatro se��es
                        vSplittedTAG = Split(sCtrlTAG, "-")
                        iTagSection = 3
    'Stop
                        
                        'Avalia a 3a se��o com par�metros de Behvr
                        '-------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------
                        
                        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                        If vSplittedTAG(iTagSection - 1) <> "" Then
'vA = vSplittedTag(2)
If gBbDepurandoLv01b Then MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Chama [ pbSub41_CtrlsBehvrDictBuild ] pra" & vbCr & "inclus�o de [ " & sTriggCtrl & " ] no dict [ dictCtrlBehvrParams(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
    
                            lngEvalTAGedCtrls = lngEvalTAGedCtrls + 1
    
                            'Chama rotina pra iniciar a montagem do dicion�rio de [ CtrlsBehvrParams ]
                            '-------------------------------------------------------------
                            On Error GoTo -1
                            bFoundParams = pbSub41_CtrlsBehvrDictBuild(vSplittedTAG(iTagSection - 1), cTriggCtrl)
                            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                            
'MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub41_CtrlsBehvrDictBuild ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
'Stop
                            
                            'Se os par�metros necess�rios tiverem sido encontrados incrementa a contagem de [ controles inclu�dos ] no dict [ ]
                            If bFoundParams Then lngCtrlsInDict = lngCtrlsInDict + 1
                            
                        End If
                        
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Controle avaliado [ " & sTriggCtrl & " ]" & vbCr & vbCr & "[ " & lngEvalFormCtrls & " ] controles no [ Form ]" & vbCr & "[ " & lngEvalTAGedCtrls & " ] controles avaliados pra [ BehvrCtrls ]"
vB = vbCr & "[ " & lngCtrlsInDict & " ] controles inclu�dos em [ dictCtrlBehvrParams(sForM) ]" & vbCr & " "
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
    
                        '----------------------------------------------
                        '-------------------------------------------------------------------------------------------------------------
                    
                    'Controle n�o TAG e n�o � inclu�do no dict [ dictCtrlBehvrParams(sForM) ]
                    Else
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "[ " & sTriggCtrl & "  ] N�O tem a TAG pra inclus�o no dict" & vbCr
vB = "[ dictCtrlBehvrParams(sForM) ]" & vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
                    
                   
                    End If
                    '----------------------------------------------
                    '-------------------------------------------------------------------------------------------------------------
                
                End If
                '----------------------------------------------
                '-------------------------------------------------------------------------------------------------------------
                
'vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Controle avaliado [ " & sTriggCtrl & " ]" & vbCr & vbCr & "[ " & lngEvalFormCtrls & " ] controles no [ Form ]" & vbCr & "[ " & lngEvalTAGedCtrls & " ] controles avaliados pra [ BehvrCtrls ]"
'vB = vbCr & "[ " & lngCtrlsInDict & " ] controles inclu�dos em [ dictCtrlBehvrParams(sForM) ]" & vbCr & " " & vbCr & " "
'MsgBox vA & vB
'If gBbDepurandoLv01b Then Stop
'Stop
        End Select

NextCtrl:
    
    Next cTriggCtrl
       
'MsgBox "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr & "Conclus�o do dict [ ResetArea ]"
If gBbDepurandoLv01c Then Stop
'Stop
                 

'Stop


'    'Teste de acesso aos valores armezanados


'    For Each vA In dictFrmResetAreas(sForM)
'        Set clObjRstAreaParams = dictFrmResetAreas(sForM)(vA)
'
'    Next vA


'    Set vA = dictTrggCtrlsInForm(sForM)

'    'Set dictFormFilterGrpsTrgts(sForM) = dictTrggCtrlsInForm(sForM)
'
'    For Each vKey In dictTrggCtrlsInForm(sForM)
'Stop
'        Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(vKey)
'
'    Next vKey



'Stop
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & "Retorna de [ pbSub31_TriggCtrlDictBuild ] e" & vbCr & "[ pbSub41_CtrlsBehvrDictBuild ] ap�s avaliar todos os controles"
vB = vbCr & vbCr & "Se controles avaliados e controles inclu�dos em [ dictCtrlBehvrParams ] forem diferentes chama [ FormStatusBar01_Bld ] e inclui o erro no Log"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop

FrM_Error_SaiR:
        
        
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr & vbCr
vB = "O form [ " & sForM & " ] tem" & vbCr & "[ TrgtCtrls ] mas nenhum [ TrggCtrl ]." & vbCr & "Chama [ FormStatusBar01_Bld ] e inclui o erro no Log."
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
    
    
    'Confirma se h� [ Grupos de Filtragem ] com [ TargtCtrls ] mas sem [ Triggers ] no [ Form ]

    On Error GoTo -1

    'vA = dictFormFilterGrpsTrgts(sForM)("01")
    
    'Se verdadeiro significa que no [ Form ] h� pelo menos um [ Grupo de Filtragem ] com seu respectivo [ TargetCtrl ]
    If IsObject(dictFormFilterGrpsTrgts(sForM)) Then
        
        'Passa por todos os [ Grupos de Filtragem ] do [ Form ]
        ' e confirma se todos t�m [ TriggCtrls ] associados
        For Each vKeyFilGrp In dictFormFilterGrpsTrgts(sForM)
            
            'Se o [ Grupo de Filtragem ] ora avaliado n�o tiver [ TriggCtrls ] informa o erro no log de carga do sistema
            If Not IsObject(dictTrgg01CtrlsInGrp(vKeyFilGrp)) Then
                
                sLoadLogWarn = "O grupo de filtragem [ " & vKeyFilGrp & " ] tem o respectivo [ TrgtCtrl ] mas n�o tem [ TriggCtrls ] associados." & vbCrLf & "N�o ser� poss�vel fazer pesquisas."
                
                On Error GoTo -1
                Call FormStatusBar01_Bld(sForM, "Targts_NoTriggers", sLoadLogWarn)
                If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                
                'vA = dictTrgg01CtrlsInGrp(vKeyFilGrp).Count
            
            End If
        
        Next vKeyFilGrp
    
    End If
    
    
    'Se a quantidade de controles inclu�dos no Dict [ dictCtrlBehvrParams(sForM) ] for diferente dos controles
    ' avaliados no Form adiciona a informa��o na StatusBar
    ' de que O form tem controles sem os par�metros de comportamento
    If lngCtrlsInDict <> lngEvalTAGedCtrls Then
        
vA = "----- pbSub30_TriggCtrlDictStartUp ---------------------------------------------" & vbCr
vB = "A quantidade de controles avaliados [ " & lngEvalTAGedCtrls & " ]   e a quantidade de" & vbCr & "controles inclu�dos em [ dictCtrlBehvrParams ] [ " & lngCtrlsInDict & " ] s�o diferentes. Chama [ FormStatusBar01_Bld ] e inclui o erro no Log"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
        
    End If
    
    
    On Error GoTo -1
    Exit Sub

FrM_ErrorHandler:
'Stop

    If Err.Number = 9 Then
        'Matriz n�o cont�m os itens esperados
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTriggCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
        'sStR2 = "A [ " & iTagSection & "a se��o ] " & " de par�metros do TriggerCtrl" & vbCr & " n�o foi localizada." & vbCr & vbCr & " Esse campo poder� se comportar de forma inesperada."
        'vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "

        'Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
        
        sLoadLogWarn = "Os seguintes controles n�o tem todos os par�metros de configura��o e portanto poder�o se comportar de forma inesperada."
        On Error GoTo -1
        Call FormStatusBar01_Bld(sForM, "CtrlMissingParams", sLoadLogWarn, sTriggCtrl)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

'Stop
        GoTo NextCtrl
        'Resume Next

    Else
        'Erro n�o previsto
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
        '
        Application.Quit
        Resume Next
    End If

End Sub

Public Sub pbSub31_TriggCtrlDictBuild(vTagSection As Variant, cTriggCtrl As Control)

    Dim vA, vB, vC, vD
    Dim vE, vF, vG, vH
    Dim vTagSectionParams As Variant
    Dim iWhere As Integer
    
    Dim sStR1 As String, sStR2 As String
    Dim sParam As String
    
    Dim dDictOuter As Dictionary, dDictInner As Dictionary
    Dim sQryField As String
    Dim sFilGrp As String
    Dim sFilGrpTag As String
    Dim iSrchWildCard As Integer
    Dim iQryFldRmvCharCpt As Integer
    Dim iSrchOnChange As Integer
    Dim sCascUpDtTrgCtrl As String
    Dim sQryFieldCptClean As String
    Dim iListboxTxtClmn As Integer
    Dim sTargtCtrlSQLselect As String
    Dim bBolClctd As Boolean
        
    Dim sTrggCtrL As String
    Dim sTrgtCtrl As String
    Dim sForM As String
    Dim sRecCntCtrl As String
    Dim cRecCnt As Control
    Dim vKey As Variant
    Dim vKeyGrp As Variant, vKeyTrgt As Variant
    Dim sLoadLogWarn As String
    Dim sSQLtablesString As String
    
    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
    sTrggCtrL = cTriggCtrl.Name
    sForM = cTriggCtrl.Parent.Name
    
'Stop '----------------------------

    '-------------------------------------------------------------------------------------------------------
    'Recupera os par�metros do controle informados na TAG
    '-------------------------------------------------------------------------------------------------------
'Stop
'MsgBox "----- pbSub31_TriggCtrlDictBuild -----------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sTrggCtrL & " ] pra inclus�o em [" & Chr(160) & "dictTrgg01CtrlsInGrp" & Chr(160) & "]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
'Stop
    
    vTagSectionParams = Split(vTagSection, ".")
    
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    sParam = "Grp"
        
        
        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "O TrggCtrl [ " & sTrggCtrL & " ] n�o est� associado a" & vbCrLf & "nenhum grupo  de filtragem e n�o poder� ser pesquisado."
        
        'Mensagem de erro a ser exibida em tela
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl:     " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = " O TriggerCtrl n�o est� associado a nenhum grupo " & vbCr & "  de filtragem e n�o poder� ser pesquisado."
        
        On Error GoTo -1
        sFilGrp = GetTagParams(sParam, vTagSectionParams, , True, "", 1, , True, sStR1, sStR2, True, "MissingTrggFilGrp", cTriggCtrl, sLoadLogWarn)
'Stop
        'sFilGrp = GetTagParams(sParam, vTagSectionParams, , "", 1, , sStR1, sStR2)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
'MsgBox "----- pbSub31_TriggCtrlDictBuild -----------------------------------------------" & vbCr & vbCr & "Avalia grupo do Trigger [ " & sTrggCtrL & " ] [ " & sFilGrp & " ]" & vbCr & " " & vbCr & " "
'If gBbDepurandoLv01c Then Stop
'Stop
        
        
        If sFilGrp = "" Then Exit Sub
        'vA = dictFormFilterGrpsTrgts(sForM)
        
        
'If IsObject(dictFormFilterGrpsTrgts(sForM)(sFilGrp)) Then
'parei aqui
        
        'dictFormFilterGrpsTrgts (sForM)
        
        If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM) = New Dictionary
        If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp) = New Dictionary
        If Not dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Exists(sTrggCtrL) Then dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Add sTrggCtrL, sFilGrp
    '-----------------------------------------------------------------------------------
    'O Grupo de Filtragem do TrggCtrl foi identificado. Continua o procedimento
    '-----------------------------------------------------------------------------------
    '-------------------------------
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & vbCr & "erro na carga do RecCnt de [ " & sTrggCtrL & " ], grupo [ " & sFilGrp & " ]" & vbCr & " " & vbCr & " "
'Stop
    'Identifica na Tag do controle, o [ Campo da Consulta ] a ser usado na filtragem
    sQryField = vTagSectionParams(1)

    If gBbEnableErrorHandler Then On Error Resume Next
    
    
    'Se [ sFilGrp ] n�o existir no dict [ dictFormFilterGrpsTrgts(sForM) ] significa que n�o h� [ TrgtCtrls ]
    ' associados a esse grupo
    ' Nesse caso desconsidera o [ TrggCtrl ] pra filtragem
    vA = dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp)
    

    '*Erro aqui
    'Ao executar a linha abaixo o [ sFilGrp ] � adicionado ao dicion�rio mesmo gerando o erro 424
    'A solu��o encontrada foi remover o item criado caso tenha gerado o erro 424, o que significa que este grupo n�o deve ser adicionado ao dicion�rio

'    Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp).Key(0)
'
'    If (Err.Number = 424) Then  'object required  Grupo informado no Trigger n�o consta no dicion�rio de [ Grupos de Filtragem ]
    
    If Not vA Then
        'If dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp) Then dictFormFilterGrpsTrgts(sForM).Remove (sFilGrp)
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = " O TriggerCtrl est� associado ao Grupo de Filtragem [ " & sFilGrp & " ]" & vbCr & "  que n�o foi carregado na inicializa��o do sistema." & vbCr & "  Esse TriggerCtrl n�o poder� ser pesquisado."
        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
        
        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
                       
        sLoadLogWarn = "O TriggerCtrl est� associado ao Grupo de Filtragem [ " & sFilGrp & " ]" & vbCrLf & "  que n�o foi carregado na inicializa��o do sistema." & vbCrLf & "  Esse TriggerCtrl n�o poder� ser pesquisado."
        On Error GoTo -1
        Call FormStatusBar01_Bld(sForM, "FilGrpError", sLoadLogWarn, sTrggCtrL)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
'Stop
        Exit Sub
    
    End If
    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
'MsgBox "varredura de trgctrls"
'Stop

    For Each vKeyTrgt In dictFormFilterGrpsTrgts(sForM)(sFilGrp)
        
        If Not IsEmpty(vKeyTrgt) And vKeyTrgt <> "" Then Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)(vKeyTrgt)
        sTrgtCtrl = clObjTargtCtrlParam.sTargtCtrlName
        
        If Not clObjTargtCtrlParam.dictQryFields.Exists(sQryField) Then
            
            vA = "O campo de tabela [ " & sQryField & " ] indicado nos par�metros" & vbCr & " do TriggerCtrl n�o foi localizado na consulta"
            vB = vbCr & " [ " & "" & " ], base de dados do TrgtCtrl." & vbCr & vbCr & " N�o ser� poss�vel filtrar por esse TriggCtrl."
            sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
            sStR2 = vA & vB
            vC = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
            
            Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vC)
            
            sLoadLogWarn = "O campo de tabela indicado no TriggCtrl n�o foi localizado." & vbNewLine & "N�o ser� poss�vel filtrar por esse campo."
            On Error GoTo -1
            Call FormStatusBar01_Bld(sForM, "QryFieldNotFound", sLoadLogWarn, sTrggCtrL)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
            
            Exit Sub
        End If
            
    Next vKeyTrgt
 
    
    
'MsgBox "----- pbSub31_TriggCtrlDictBuild -----------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sTrggCtrL & " ] pra inclus�o em" & vbCr & "[ dictTrgg01CtrlsInGrp ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
'Stop
    '-----------------------------------------------------------------------------------
    'O Campo da Consulta do TrggCtrl foi identificado com sucesso. Continua o procedimento
    '-----------------------------------------------------------------------------------
    'Identifica o par�metro WdCrd do TriggCtrl
    sParam = "WdCrd"
        'Mensagem de erro a ser inclu�da no Log de carga
        'sLoadLogWarn = "O TargetCtrl [ " & sTrgtCtrL & " ] n�o est� associado a" & vbCrLf & "nenhum grupo  de filtragem e n�o poder� ser pesquisado."
        
'parei aqui: checar por que a mensagem de alerta foi comentada
        'Mensagem de erro a ser exibida em tela
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        'vA = " Op��es: busca; busca*; *busca*" & vbCr & vbCr & " Esse TriggerCtrl poder� gerar resultados inesperados " & vbCr & " de filtragem."
        'sStR2 = "O par�metro [ " & sParam & " ] indicando como esse campo" & vbCr & " dever� ser filtrado n�o foi informado." & vbCr & vA
        
        On Error GoTo -1
        iSrchWildCard = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, 2, , , , True, "MissingWdCrd", cTriggCtrl)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
        'iSrchWildCard = GetTagParams(sParam, vTagSectionParams, , 0, 0, 2, sStR1, sStR2)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
    '--------------------------------------------------------------------------------------------
    'Nesse trecho se houver erro de falta de par�metros � assumido o valor padr�o
    ' portanto n�o � necess�rio gerar mensagem de erro
    '--------------------------------------------------------------------------------------------
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "erro na carga do RecCnt de [ " & sTrggCtrL & " ]"
'Stop
    
    'A exist�ncia do par�metro [ iQryFldRmvCharCpt ] N�O PRECISA SER TESTADA pois
    ' se estiver vazio ser� considerado como Zero
    sParam = "Rmv"
        'Mensagem de erro a ser inclu�da no Log de carga
        'sLoadLogWarn = ""
        
        On Error GoTo -1
        iQryFldRmvCharCpt = GetTagParams(sParam, vTagSectionParams, , False, 0, 0)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

    'Identifica o texto a ser exibido no controle de contagem de Registros
    vA = vTagSectionParams(1)
    If iQryFldRmvCharCpt > 0 Then
        vB = Len(sQryField) - iQryFldRmvCharCpt
        sQryFieldCptClean = Mid(vA, 1, vB)

    Else
        sQryFieldCptClean = vA

    End If
            
            
    'A exist�ncia do par�metro [ SrchOnChange ] N�O PRECISA SER TESTADA pois
    ' se estiver vazio ser� considerado como Zero
    sParam = "SrchOnChg"
        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "O par�metro [ " & sParam & " ] dos Controles a seguir n�o foi configurado com uma op��o v�lida." & vbCrLf & "Os controles poder�o n�o se comportar como esperado."
        
        'Mensagem de erro a ser exibida em tela
        
        On Error GoTo -1
        iSrchOnChange = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, 1, , , , True, "MissingSrchOnChge", cTriggCtrl, sLoadLogWarn)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

    'A exist�ncia do par�metro [ CascUpDt ] indicando um eventual TargtCtrl secund�rio
    ' que deva ser atualizado em cascata N�O PRECISA SER TESTADA pois se estiver vazio
    ' ser� considerado como VAZIO e n�o ir� disparar uma filtragem em cascata
    sParam = "CascUpDt>"
        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "O controle indicado como CascadUpDt [ " & sTrggCtrL & " ] n�o foi localizado."
        
        'Mensagem de erro a ser exibida em tela
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = " O controle indicado como CascadUpDt n�o foi localizado."
        
        On Error GoTo -1
        sCascUpDtTrgCtrl = GetTagParams(sParam, vTagSectionParams, cTriggCtrl, False, "", , , True, sStR1, sStR2, True, "CascdNotFound", cTriggCtrl, sLoadLogWarn)
        'sCascUpDtTrgCtrl = GetTagParams(sParam, vTagSectionParams, cTriggCtrl, "", , , sStR1, sStR2)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

'Stop
    vA = cTriggCtrl.Name
    'Se o controle for um Listbox ou Combobox reativa o tratamento de erro
    If cTriggCtrl.ControlType = 110 Or cTriggCtrl.ControlType = 111 Then
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
'If gBbDepurandoLv01b Then MsgBox ".TxtClmn --------------------------------------------------------------------------"
'Stop
        'Se o Texto de Exibi��o do controle estiver na 2a coluna da tabela fonte
        ' a exist�ncia do par�metro [ TxtClmn ] N�O PRECISA SER TESTADA pois se estiver
        ' vazio ser� considerado como UM, ou seja, a 2a coluna
        sParam = "TxtClmn"
            'Mensagem de erro a ser inclu�da no Log de carga
            sLoadLogWarn = "O par�metro [ " & sParam & " ] dos Controles a seguir n�o foram definidos." & vbCrLf & "A filtragem por esses campos poder� gerar resultados inesperados."
            
            'Mensagem de erro a ser exibida em tela
            'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
            'sStR2 = "O par�metro [ TxtClmn ] indicando a String a ser usada pra" & vbCr & " filtragem desse campo n�o foi definida." & vbCr & vbCr & " Esse TriggerCtrl poder� gerar resultados inesperados " & vbCr & " de filtragem."
            
            On Error GoTo -1
            iListboxTxtClmn = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, , , , , True, "MissingTxtClmn", cTriggCtrl, sLoadLogWarn)
            
            'iListboxTxtClmn = GetTagParams(sParam, vTagSectionParams, , 1, 0, , sStR1, sStR2)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

    End If
    
    'A partir daqui o tratamento de erro deve ser retomado
    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
'End If
    
'Stop
    
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    'Foram identificadas nos par�metros do [ TrggCtrl ] as informa��es necess�rias
    ' prossegue com a montagem dos dicion�rios [ dictTrgg00GrpsInForm ] e [ dictTrgg01CtrlsInGrp ]
    ' atrav�s dos par�metros dos [ TrggCtrls ] da classe cls_02aTrggCtrlParams
    ' agrupados por Formul�rio e por Grupo de filtragem
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    
    'vA = dictTrgg00GrpsInForm("Form01")("01")("Ctrl 01C")
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "dictTrggCtrlsInForm(sForM) create"
'Stop
    '-----------------------------
    '-------------------------------------------------------
    'Cria o dicion�rio com os Trigg Controls de cada formul�rio, indicando o Grupo de Filtragem associado a cada um
    If Not IsObject(dictTrggCtrlsInForm(sForM)) Then Set dictTrggCtrlsInForm(sForM) = New Dictionary
    'Set dDicT = dictTrggCtrlsInForm(sForM)

'Stop
    'Se o controle j� foi inclu�do no dicion�rio
    If dictTrggCtrlsInForm(sForM).Exists(sTrggCtrL) = True Then
        Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(sTrggCtrL)
    
    
    Else
       'Cria um novo objeto [ clObjTriggCtrlParam ] da Classe [ cls_02aTrggCtrlParams ] pra ser inclu�do no Dict
        Set clObjFilGrpsByForm = New cls_03aCtrlsGrpsByForm
        dictTrggCtrlsInForm(sForM).Add sTrggCtrL, clObjFilGrpsByForm

        clObjFilGrpsByForm.sCtrlName = sTrggCtrL
        'Set clObjFilGrpsByForm.cCtrl = cTriggCtrl
        clObjFilGrpsByForm.sFilGrp = sFilGrp
        
    End If
    '-------------------------------------------------------
    '-----------------------------
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Add Outer dict"
'Stop
    
    '-----------------------------
    '-------------------------------------------------------
    'Cria o dicion�rio com grupos de filtragem pro formul�rio corrente, caso ele ainda n�o tenha sido criado
    If Not IsObject(dictTrgg00GrpsInForm(sForM)) Then Set dictTrgg00GrpsInForm(sForM) = New Dictionary
    Set dDictOuter = dictTrgg00GrpsInForm(sForM)
    
        
        'Cria o dicion�rio com os [ TriggCtrls ] associados ao Grupo de Filtragem ora avaliado, caso ele ainda n�o tenha sido criado
        If Not IsObject(dictTrgg01CtrlsInGrp(sFilGrp)) Then Set dictTrgg01CtrlsInGrp(sFilGrp) = New Dictionary
        Set dDictInner = dictTrgg01CtrlsInGrp(sFilGrp)
        
        'Set dictTrgg01CtrlsInGrp("02") = New Dictionary

        'Se o controle j� foi inclu�do no dicion�rio
        If dDictInner.Exists(sTrggCtrL) = True Then
            Set clObjTriggCtrlParam = dDictInner(sTrggCtrL)
        
        Else
'Stop
           'Cria um novo objeto [ clObjTriggCtrlParam ] da Classe [ cls_02aTrggCtrlParams ] pra ser inclu�do no [ dDictInner ]
            Set clObjTriggCtrlParam = New cls_02aTrggCtrlParams
            dDictInner.Add sTrggCtrL, clObjTriggCtrlParam
            
            'vA = sQryFieldCptClean
'Stop
            'Set clObjTriggCtrlParam.cCtrl = cTriggCtrl
            With clObjTriggCtrlParam
                .sCtrlName = sTrggCtrL
                .sQryField = sQryField
                .bBolClctd = bBolClctd
                .sFilGrp = sFilGrp
                .sQryFieldCptClean = sQryFieldCptClean
                .iSrchWildCard = iSrchWildCard
                .iSrchOnChange = iSrchOnChange
                .sCascUpDtTrgCtrl = sCascUpDtTrgCtrl
                .iListboxTxtClmn = iListboxTxtClmn
    '           clObjTriggCtrlParam.iClctdStrSze = iClctdStrSze
    '           clObjTriggCtrlParam.iBdCln=
    
            End With

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Inner Dict add"
'Stop
            
            Set dDictOuter = dictTrgg00GrpsInForm(sForM)
           ' dDictOuter (sFilGrp)
            'Se o Dict com os controles do Grupo j� foi inclu�do no Dict do Form
            If dDictOuter.Exists(sFilGrp) = False Then
                dDictOuter.Add sFilGrp, dDictInner
            
            End If
            
        End If
    '-------------------------------------------------------
    '-----------------------------

'Stop '-----------------------
    'Teste de acesso aos valores armezanados
'    Set dDict = dictTrggCtrlsInForm(sForM)

'    'Acesso aos valores armazenados no dict [ dictGetListSrchVals ] do objeto [ clObjTriggCtrlParam. ] dentro de [ dictTrgg01CtrlsInGrp(sFilGrp) ]
'    For Each vKey In dictTrgg01CtrlsInGrp(sFilGrp)
''Stop
'        Set clObjTriggCtrlParam = dictTrgg01CtrlsInGrp(sFilGrp)(vKey)
'        'vA = clObjTriggCtrlParam.dictGetListSrchVals(2)
'
'
'    Next vKey
'
'
'    For Each vKey In dictTrggCtrlsInForm(sForM)
''Stop
'        Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(vKey)
'
'    Next vKey

    
'    Set dDict = dictTrgg00GrpsInForm(sForm)
'    Set dDict = dictTrgg01CtrlsInGrp("01")
'    Set dDict = dictTrgg01CtrlsInGrp("03")
'
'    For Each vKeyGrp In dictTrgg00GrpsInForm(sForm)
'        vA = vKeyGrp
'
'        For Each vKeyTrgt In dictTrgg01CtrlsInGrp(vKeyGrp)
'            vB = vKeyTrgt
'            Set clObjTriggCtrlParam = dictTrgg01CtrlsInGrp(vKeyGrp)(vKeyTrgt)
'            vC = clObjTriggCtrlParam.sQryField
'
'        Next vKeyTrgt
'
'    Next vKeyGrp

'Stop
    
FrM_Error_SaiR:
    On Error GoTo -1
    Exit Sub

FrM_ErrorHandler:
'Stop
    
'    If Err.Number = 9 Then
'        'Matriz n�o cont�m os itens esperados
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = "O par�mentro " & " [ " & sParam & " ] " & " do TriggerCtrl n�o foi localizado." & vbCr & " Esse campo ser� desconsiderado para filtragem."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'        Exit Sub
'
'    ElseIf Err.Number = 2465 Then
'        'Controle do formul�rio n�o foi localizado
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "Listbox: " & "       [ " & sTrgtCtrL & " ]" & vbCr & "Contr. de Contag. de Regs: " & " [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = " O controle de contagem de registros n�o foi localizado." & vbCr & "  N�o ser� poss�vel exibir a contagem de registros associada."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'        sRecCntCtrl = ""
'        Resume Next
'
'    Else
        'Erro n�o previsto
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
        Resume Next
        Application.Quit

'    End If


End Sub


Public Sub pbSub32_TriggCtrlDictBuild(vTagSection As Variant, cTriggCtrl As Control)

   '-------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------
    'Tendo confirmado, no trecho anterior, a exist�ncia do [ Grupo de Filtragem ] no formul�rio
    ' inicia a varredura pelos [ TrgtCtrls ] de [ sFilGrp ] pra identificar os controles que devem ser pesquisados
    For Each vKeyTrgt In dictFormFilterGrpsTrgts(sForM)(sFilGrp)
    
        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)(vKeyTrgt)
        
        sTrgtCtrl = clObjTargtCtrlParam.sTargtCtrlName
    
    
    
    'If gBbDepurandoLv01b Then MsgBox "Teste - Confirma se o campo de consulta informado na TAG do TriggerCtrl [ " & sTrggCtrL & " ] existe"
    'Stop
    
    
        '----------------------------------------------------------------------------------------------
        'Checar se o campo [ sQryField ] indicado no TriggCtrl [ sTrggCtrL ] existe no grid da consulta do TargtCtrl [ sTrgtCtrl ]
        On Error GoTo -1
        NstdVarQryFld = GetFldInQryGrid(sForM, sTrgtCtrl, sQryField)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        '-----------------------------------------------------------------------------------
        
        '----------------------------------------------------------------------------------------------
        'Se o campo n�o for encontrado no grid da consulta, procura em todas as tabelas e Queries usadas na consulta
        If Not NstdVarQryFld.bFoundQryFld Then
        
            '-----------------------------------------------------------------------------------
            'Checar se o campo [ sQryField ] indicado no TriggCtrl [ sTrggCtrL ] existe em alguma das tabelas
            ' da consulta do TargtCtrl [ sTrgtCtrl ]
            ' se n�o existir exibe alerta e remove o Controle do dicion�rio do Grupo de Filtragem
            sSQLtablesString = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT & " " & clObjTargtCtrlParam.sClsLstbxSQL_bFROM
            On Error GoTo -1
    MsgBox ""
    Stop
            NstdVarQryFld = GetFldInQryGridTbls(sForM, sTrgtCtrl, sSQLtablesString, sQryField)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
        End If
        '----------------------------------------------------------------------------------------------
        
    
    
    'Stop
        'Se o campo da consulta n�o tiver sido encontrado na tabela de dados exibe o alerta
        ' e n�o inclui o controle no dicion�rio [ dictTrgg01CtrlsInGrp(sFilGrp) ]
        If Not NstdVarQryFld.bFoundQryFld Then
    
            vA = "O campo de tabela [ " & sQryField & " ] indicado nos par�metros" & vbCr & " do TriggerCtrl n�o foi localizado na consulta"
            vB = vbCr & " [ " & NstdVarQryFld.sQry & " ], base de dados do TrgtCtrl." & vbCr & vbCr & " N�o ser� poss�vel filtrar por esse TriggCtrl."
            sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrl & " ]" & vbCr & "-------------------------------------------------------------------------------"
            sStR2 = vA & vB
            vC = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
    
            Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vC)
    
            sLoadLogWarn = "O campo de tabela indicado no TriggCtrl n�o foi localizado." & vbNewLine & "N�o ser� poss�vel filtrar por esse campo."
            On Error GoTo -1
            Call FormStatusBar01_Bld(sForM, "QryFieldNotFound", sLoadLogWarn, sTrggCtrL)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    'Stop
            Exit Sub
    Stop
        End If
        '-----------------------------------------------------------------------------------
        
        'Se o campo da consulta n�o tiver sido informado suspende a carga do controle no sistema
        If sQryField = "" Then
            sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
            sStR2 = "O TriggerCtrl n�o possui a indica��o do campo de tabela" & vbCr & " pra consultas e n�o poder� ser pesquisado."
            vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
            
            Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
            
            sLoadLogWarn = "O TriggerCtrl n�o possui a indica��o do campo de tabela" & vbNewLine & " pra consultas e n�o poder� ser pesquisado."
            On Error GoTo -1
            Call FormStatusBar01_Bld(sForM, "MissingQryField", sLoadLogWarn, sTrggCtrL)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    Stop
            Exit Sub
    'Stop
        Else
            '-----------------------------------------------------------------------------------
            'Avalia se deve ser usado o Nome de Campo [ sQryField ] informado ou o respectivo Campo Calculado
            '-----------------------------------------------------------------------------------
    'Stop
            If gBbEnableErrorHandler Then On Error Resume Next
            'Recupera, no dict [ dictFormFilterGrpsTrgts ] o SQL Select do TargtCtrl associado ao Grupo de Filtragem ora avaliado
            
            'Se [ sFilGrp ] n�o existir no dict [ dictFormFilterGrpsTrgts(sForM) ] significa que n�o h� [ TrgtCtrls ]
            ' associados a esse grupo
            ' Nesse caso desconsidera o [ TrggCtrl ] pra filtragem
            vA = dictFormFilterGrpsTrgts(sForM).Exists(sFilGrp)
            
            If Not vA Then
    '        Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)
    '
    '        'Se o [ FiltGrp ] informado para o [ TrggCtrl ] n�o constar no dicion�rio de Grupos de Filtragem
    '        ' significa que o Grupo n�o existe ou n�o foi carregado corretamente o que indica
    '        ' que o [ Grupo de Filtragem ] n�o tem um TargtCtrl tamb�m associado
    '        ' Nesse caso desconsidera o [ TrggCtrl ] pra filtragem
    '        If (Err.Number = 424) Then  'object required  Grupo informado no Trigger n�o consta no dicion�rio de [ Grupos de Filtragem ]
    'Stop
                sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
                sStR2 = " O TriggerCtrl est� associado ao Grupo de Filtragem [ " & sFilGrp & " ]" & vbCr & "  que n�o foi carregado na inicializa��o do sistema." & vbCr & "  Esse TriggerCtrl n�o poder� ser pesquisado."
                vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
                
                Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
                
                sLoadLogWarn = "O TriggerCtrl est� associado ao Grupo de Filtragem [ " & sFilGrp & " ]" & vbCrLf & "  que n�o foi carregado na inicializa��o do sistema." & vbCrLf & "  Esse TriggerCtrl n�o poder� ser pesquisado."
                On Error GoTo -1
                Call FormStatusBar01_Bld(sForM, "FilGrpError", sLoadLogWarn, sTrggCtrL)
                If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                
    Stop
                Exit Sub
            
            End If
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
            
            'N�o foi localizado um [ TargtCtrl ] associado ao grupo de filtragem do [ TriggCtrl ]
            ' quando poss�vel confirmar quando ocorre esse erro
            If (Err.Number = 13) Then  'type mismatch
                sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "Grupo de Filtragem: [ " & sFilGrp & " ]" & vbCr & "-------------------------------------------------------------------------------"
                sStR2 = " N�o foi localizado no Formul�rio nenhum TargtCtrl" & vbCr & "  associado ao Grupo de Filtragem." & vbCr & "  N�o ser� poss�vel filtrar por esse TriggCtrl. Checar Erro --"
                vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
                Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
    Stop
                Exit Sub
            
            End If
            
            
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
            
            sTargtCtrlSQLselect = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT
            
            '-----------------------------------------------------------------------------------
            'Inicia a avalia��o do campo calculado
            '-----------------------------------------------------------------------------------
            
            'Verifica se o sQryField informado no par�metro est� associado � express�o "AS" no SQL
            ' Isso indica que � um campo calculado o que exije um tratamento diferenciado
            ' para ser recuperado
            vA = " AS " & sQryField
            iWhere = InStr(sTargtCtrlSQLselect, vA)
    
    'If gBbDepurandoLv01b Then MsgBox "----- pbSub31_TriggCtrlDictBuild -----------------------------------------------" & vbCr & "teste " & sTrggCtrL
    'stop
            If iWhere > 0 Then
                'Chama a fun��o pra montar o Campo Calculado
                ' O nome de campo de filtragem [ sQryField ] identificado nos par�metros do controle
                ' � substitu�do pelo respectivo campo calculado
                vB = sGetClcltdField(sTargtCtrlSQLselect, " AS " & sQryField)
    'Stop
                sQryField = vB
                
                If vC = "SELECT NotFound" Then
                    sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "TriggerCtrl: " & "  [ " & sTrggCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
                    vA = vbCr & "  localizar o trecho do SELECT, impedindo a montagem da" & vbCr & "  f�rmula associada ao campo pra uso no SQL." & vbCr & "  O TriggerCtrl n�o poder� ser pesquisado."
                    sStR2 = " � poss�vel que o TriggerCtrl seja um campo calculado," & vbCr & "  mas os par�metros informados no controle n�o permitiram" & vA
                    vB = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
                    Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vB)
    Stop
                    Exit Sub
                    
                End If
                
                bBolClctd = True
                
            Else
                bBolClctd = False
            
            End If
        
        End If
        '-------------------------------
        '-----------------------------------------------------------------------------------
    
    Next vKeyTrgt
    'Encerra a varredura pelos [ TrgtCtrls ] de [ sFilGrp ]
    '-------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------
    
End Sub



Public Function pbSub41_CtrlsBehvrDictBuild(vTagSection As Variant, cCtrL As Control)

    Dim vA, vB, vC
    Dim sForM As String
    Dim sCtrL As String
    Dim sParam As String
    Dim iWhere As Integer
    Dim vTagSectionParams As Variant
    Dim sStR1 As String, sStR2 As String
    Dim sHLclr As String
    Dim sOnDrty As String
    Dim sMskd As String
    Dim bMskdCtrlEventFound  As Boolean
    Dim sLoadLogWarn As String
    Dim sTmP As String
    Dim sModName As String, sSubName As String, sSearchTerm As String
    
    If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    
    sForM = cCtrL.Parent.Name
    sCtrL = cCtrL.Name
    
If gBbDepurandoLv01c Then MsgBox "----- pbSub41_CtrlsBehvrDictBuild ----------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sCtrL & " ] pra inclus�o" & vbCr & "em [ dictCtrlBehvrParams(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
    
    '-------------------------------------------------------------------------------------------------------
    'Recupera os par�metros do controle armazenados na TAG
    '-------------------------------------------------------------------------------------------------------
    
    vTagSectionParams = Split(vTagSection, ".")
'Stop
    
    'Confirma se [ sCtrL ] ora avaliado possui o evento [ Change ] com a chamada pra rotina [ MskdTxtbox02_TextMask ]
    sModName = "Form_" & sForM ' "frm_01(1)cProdEstoque"
    sSubName = sCtrL & "_Change"
    sSearchTerm = "MskdTxtbox02_TextMask"
    
    bMskdCtrlEventFound = FindCodeLineInSub(sModName, sSubName, sSearchTerm)
'Stop

    
    
    'Verifica se foi identificado o par�metro [ HLclr ] do controle indicando se ele deve mudar de cor no foco
    sParam = "HLclr"
        
        On Error GoTo -1
        sHLclr = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, 1, , , , True, "MissingHLclr", cCtrL)

    sParam = "OnDrty"
        
'parei aqui: checar por que a mensagem de alerta foi comentada
        'Mensagem de erro a ser inclu�da no Log de carga
        'sLoadLogWarn = ""
        
        'Mensagem de erro a ser exibida em tela
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: " & "     [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------" & vbCr
        'sStR2 = "O par�metro [ " & sParam & " ] do Controle n�o foi" & vbCr & " configurado com uma op��o v�lida." & vbCr & vbCr & " O controle poder� n�o se comportar como esperado."
        
        On Error GoTo -1
        sOnDrty = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, 1, , , , True, "MissingOnDrty", cCtrL)
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler

        
        
    'Se o controle estiver configurado como [ bOnDirty ] TRUE
    ' confirma se ele � um [ Trigger ]
    If sOnDrty = 1 Then
    
If gBbDepurandoLv01c Then MsgBox "----- pbSub41_CtrlsBehvrDictBuild ----------------------------------------------" & vbCr & vbCr & "Inclui [ " & sCtrL & " ]  no Log de erros de carga do sistema caso " & vbCr & "esteja configurado como [ bOnDirty ] mas n�o seja um [ Trigger ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
    
        'Confirma se [ sCtrL ] � um [ Trigger ]
        'Verifica se o dict [dictTrggCtrlsInForm(sForM)] foi criado, o que indica que h� [ TrggCtrls ] carregados
        vA = IsObject(dictTrggCtrlsInForm(sForM))
        
        'Se o dicion�rio de [ TrggCtrls ] n�o existir ou se ele existir mas [ sCtrL ] n�o tiver sido inclu�do, indica que ele N�O � um trigger
        If vA Then vB = dictTrggCtrlsInForm(sForM).Exists(sCtrL) Else vB = False
    
        
        'Se o controle N�O for um [ TrggCtrl ] e foi identificado o par�metro [ OnDrty ] TRUE, muda o par�metro pra FALSE
        ' e carrega, na StatusBar, a informa��o de par�metro [ OnDirty ] definido como True em controle N�O Trigger
        If Not vB Then 'n�o � trigger
            sOnDrty = 0
            
            'Chama rotina pra incluir na [ StatusBar ] do [ Form ] a informa��o de que
            ' h� controles que n�o s�o [ TrggCtrl ] marcados pra Destacar o valor preenchido
            sLoadLogWarn = "H� controles configurados com [ OnDirty ] mas N�O carregados como [ Trigger ]." & vbNewLine & "Esses controles n�o ir�o mudar de cor no ""sujo""."
    
    If gBbDebugOn Then Debug.Print sLoadLogWarn
vA = "----- pbSub41_CtrlsBehvrDictBuild ----------------------------------------------" & vbCr & vbCr & "Chama [ FormStatusBar01_Bld ] pra incluir [ " & sCtrL & " ] em"
vB = vbCr & "[ dictFormsParams(sForM).clObjFormsParams.dForm_StatusBarWarns ]" & vbCr & "indicando erro na carga do controle"
If gBbDepurandoLv01c Then MsgBox vA & vB
If gBbDepurandoLv01c Then Stop
'Stop
            On Error GoTo -1
            Call FormStatusBar01_Bld(sForM, "DirtyTrue_NoTrgg", sLoadLogWarn, sCtrL)
            If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
        End If
'Stop
    End If
'Stop

    'Verifica se foi identificado o par�metro [ Mskd ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    '------------------------------------------------------------------------------------------------------------------------
    sParam = "Mskd"
        
'parei aqui: checar por que a mensagem de alerta foi comentada
        'Mensagem de erro a ser inclu�da no Log de carga
        'sLoadLogWarn = "O TargetCtrl [ " & sTrgtCtrL & " ] n�o est� associado a" & vbCrLf & "nenhum grupo  de filtragem e n�o poder� ser pesquisado."
        
        'Mensagem de erro a ser exibida em tela
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: " & "     [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------" & vbCr
        'sStR2 = "O par�metro [ " & sParam & " ] do Controle n�o foi" & vbCr & " configurado com uma op��o v�lida." & vbCr & vbCr & " O controle poder� n�o se comportar como esperado."
        
        On Error GoTo -1
        sMskd = GetTagParams(sParam, vTagSectionParams, , False, 0, 0, 1, , , , True, "MissingMskd", cCtrL)
        'sMskd = GetTagParams(sParam, vTagSectionParams, , 0, , 1, sStR1, sStR2, cCtrL, "MskdParamError")
        If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
    '------------------------------------------------------------------------------------------------------------------------
    
    
If gBbDepurandoLv01c Then MsgBox "----- pbSub41_CtrlsBehvrDictBuild ----------------------------------------------" & vbCr & vbCr & "Salva os par�metros de [ " & sCtrL & " ] em" & vbCr & "[ dictCtrlBehvrParams(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop

    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------
    'Cria uma nova varia��o do dicion�rio pro Formul�rio corrente, caso ele ainda n�o tenha sido criado,
    ' pra depois carregar os par�metros de comportamento dos Controles
    '-------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------
    If Not IsObject(dictCtrlBehvrParams(sForM)) Then Set dictCtrlBehvrParams(sForM) = New Dictionary

    'Verifica se o Controle j� foi inclu�do no [ Dict ]
    If dictCtrlBehvrParams(sForM).Exists(sCtrL) = True Then
        Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sCtrL)
'Stop
    Else
        'Cria um novo objeto [ clObjCtrlBehvrParams ] da Classe [ cls_11aCtrlBehvrParams ] pra ser inclu�do no Dict
        Set clObjCtrlBehvrParams = New cls_11aCtrlBehvrParams

        'Adiciona um novo item no dicion�rio [ dictFormFilterGrpsTrgts ] e guarda nele o objeto [ clObjCtrlBehvrParams ]
        ' com os respectivos par�metros do Grupo de Filtragem definidos na classe [ cls_11aCtrlBehvrParams ]
        dictCtrlBehvrParams(sForM).Add sCtrL, clObjCtrlBehvrParams
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "build [ " & sCtrL & " ]" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & ""
'Stop
        With clObjCtrlBehvrParams
            'Atribui ao controle os par�metros esperados pela Classe [ cls_11aCtrlBehvrParams ]
            .sCtrlName = sCtrL
            .bColorHighlight = IIf(sHLclr = 0, False, True)
            .bOnDirty = IIf(sOnDrty = 0, False, True)
            .bMskdCtrl = IIf(sMskd = 0, False, True)
            .bMskdCtrlEventFound = bMskdCtrlEventFound
'Stop
        End With
    
    End If



'Stop
    'Confirma se todos os par�metros do Controle foram encontrados
    If UBound(vTagSectionParams) < 3 Then
        sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: " & "     [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        sStR2 = "A se��o [ " & "BehvrParams" & " ] da TAG do Controle n�o tem" & vbCr & " todos os par�metros necess�rios." & vbCr & vbCr & " O controle poder� N�O se comportar como esperado."
    
        vB = " Erro [ " & Err.Number & " ] "
           
        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vB)
Stop
    End If


    'Se tiver passado pela rotina inteira significa que conseguiu carregar os par�metros
    ' nesse caso devolve TRUE
    pbSub41_CtrlsBehvrDictBuild = True

FrM_Error_SaiR:
    On Error GoTo -1
    Exit Function

FrM_ErrorHandler:
Stop
'
'    If Err.Number = 9 Then
'        'Matriz n�o cont�m os itens esperados
'        sStR1 = "Formul�rio:  [ " & sForm & " ]" & vbCr & "Controle: " & "    [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
'        sStR2 = "O par�mentro" & " [ " & sParam & " ]" & " do controle n�o foi localizado." & vbCr & " O Controle poder� n�o se comportar como esperado."
'        vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
'        Exit Function
'
'    Else
        'Exibe o c�digo de erro
        MsgBox Err.Description, , "Erro:" & Err.Number

        'Avisa ao usu�rio que o sistema ser� encerrado pois ocorreu um erro n�o previsto em c�digo
        sStR1 = "-------------------------------------------------------------------------------" & vbCr & " Erro de sistema n�o previsto."
        sStR2 = "O sistema ser� encerrado!"

        Call msgboxErrorAlert(sStR1, sStR2, vbExclamation)
Stop
Resume Next
        Application.Quit

'    End If

    Resume Next
    
End Function


Public Sub pbSub51_RstAreaDictBuild(vTagSection As Variant, cRstAreaCtrL As Control)

    Dim vA, vB, vC
    Dim vTagSectionParams As Variant 'vSplittedTag As Variant
    Dim sParam As String
    Dim sCtrL As String
    Dim sForM As String
    Dim sStR1 As String, sStR2 As String
    Dim sRstAr As String
    Dim vFilGrp As Variant
    Dim sFilGrp As String
    Dim sLoadLogWarn As String
    
   'Localiza dentro da TAG do Controle qual a �rea de reset que deve ser tratada
    'iWhere = InStr(1, vTAGsectionParams(0), "TrgtCtrl")
    sForM = cRstAreaCtrL.Parent.Name
    sCtrL = cRstAreaCtrL.Name
                                  
If gBbDepurandoLv01c Then MsgBox "----- pbSub51_RstAreaDictBuild -------------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sCtrL & " ] pra inclus�o" & vbCr & "em [ dictFrmResetAreas(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop

    'Separa [ vTagSection ] em matriz de �nico elemento, com o par�metro
    vTagSectionParams = Split(vTagSection, ".")
    
    
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    sParam = "RstAr"

'parei aqui: checar por que a mensagem de alerta foi comentada

        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "O par�metro [ " & sParam & " ] dos Controles a seguir n�o foi informado" & vbCrLf & "Esses controles n�o poder�o ser esvaziados com bot�es [ Reset]."
        
        'Mensagem de erro a ser exibida em tela
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: " & "     [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------" & vbCr
        'sStR2 = "O par�metro [ " & sParam & " ] do Controle n�o foi" & vbCr & " configurado com uma op��o v�lida." & vbCr & vbCr & " O controle n�o poder� ser esvaziado com bot�es [ Reset ]."

        sRstAr = GetTagParams(sParam, vTagSectionParams, , False, "", 0, , , , , True, "MissingRstAr", cRstAreaCtrL, sLoadLogWarn)
        'sRstAr = GetTagParams(sParam, vTagSectionParams, , "", 1, , sStR1, sStR2, cRstAreaCtrL, "RstAreaParamError")
        'Se retornar vazio n�o carrega o [ controle ] no [ dictFrmResetAreas(sForM) ]
        If sRstAr = "" Then Exit Sub
        
'----------------
'dictFrmResetAreas(sForM).RemoveAll
'----------------
        
    '-------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------
    'Cria uma nova varia��o do dict [ dictFrmResetAreas ] pro Formul�rio corrente, caso ele ainda n�o tenha sido criado,
    ' pra guardar os dados de Areas de Reset do Form
    '-------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------
    If Not IsObject(dictFrmResetAreas(sForM)) Then Set dictFrmResetAreas(sForM) = New Dictionary
'Stop
    
    'Adiciona [ sRstAr ] em [ dictFrmResetAreas(sForM) ]
    If dictFrmResetAreas(sForM).Exists(sRstAr) = True Then
        Set clObjRstAreaParams = dictFrmResetAreas(sForM)(sRstAr)
'    dictFrmResetAreas(sForM).clObjRstAreaParams
    Else
       'Cria um novo objeto [ clObjTriggCtrlParam ] da Classe [ cls_02aTrggCtrlParams ] pra ser inclu�do no Dict
        Set clObjRstAreaParams = New cls_05aResetAreasParams
        
        dictFrmResetAreas(sForM).Add sRstAr, clObjRstAreaParams
        clObjRstAreaParams.sRstAr = sRstAr
    
    End If
    
'Stop
    
    '-------------------------------------------------------------------------------------------------------------------
    'Ap�s adicionar [ sRstAr ] cria os dois dicion�rios internos
    '-------------------------------------------------------------------------------------------------------------------
    If dictFrmResetAreas(sForM)(sRstAr).dictRstArCtrls Is Nothing Then Set dictFrmResetAreas(sForM)(sRstAr).dictRstArCtrls = New Dictionary
    If dictFrmResetAreas(sForM)(sRstAr).dictRstArFilGrps Is Nothing Then Set dictFrmResetAreas(sForM)(sRstAr).dictRstArFilGrps = New Dictionary
'Stop
    '-----------------------------------------------------------------------------------------------------------------


If gBbDepurandoLv01c Then MsgBox "----- pbSub51_RstAreaDictBuild -------------------------------------------------" & vbCr & vbCr & "Avalia [ " & sCtrL & " ] " & vbCr & "ResetArea [ " & sRstAr & " ] " & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
'Stop
    
    
    'Adiciona [ sFilGrp ] em [ dictRstArFilGrps ], dicion�rio interno � classe [ clObjRstAreaParams ],
    '-----------------------------------------------------------------------------------------------------------------
    'Recupera o [ Grupo de Filragem ] de [ sCtrL ]
    '-----------------------------------------------------------------------------------------------------------------
    'Confirma se o dicion�rio de [ Triggers ] j� foi criado.
    ' caso negativo significa que o controle ora avaliado n�o � um [ Trigger ]
    If IsObject(dictTrggCtrlsInForm(sForM)) Then
        
        If dictTrggCtrlsInForm(sForM).Exists(sCtrL) Then
'Stop
            Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(sCtrL)
            sFilGrp = clObjFilGrpsByForm.sFilGrp
            
            If Not clObjRstAreaParams.dictRstArFilGrps.Exists(sFilGrp) Then
                'Set clObjLckdStatusParam = clObjRstAreaParams.dictRstArFilGrps(sFilGrp)
            
            'Else
                'Adiciona o Grupo a [ clObjRstsFilGrp.dictRstArFilGrps(sFilGrp) ]
                'Set clObjLckdStatusParam = New cls_07bLckdStatusParams
                clObjRstAreaParams.dictRstArFilGrps.Add sFilGrp, sFilGrp
        
            End If
            
        End If
        
    End If
    '-----------------------------------------------------------------------------------------------------------------
    '-------------------------------------------

If gBbDepurandoLv01c Then MsgBox "----- pbSub51_RstAreaDictBuild -------------------------------------------------" & vbCr & vbCr & "Adiciona os itens ao dicion�rio" & vbCr & " Ctrl [ " & sCtrL & " ] " & vbCr & " Grp [ " & sFilGrp & " ] " & vbCr & " RstAr" & " [ " & sRstAr & " ]"
If gBbDepurandoLv01c Then Stop
'Stop
    
    '-----------------------------------------------------------------------------------------------------------------
    'Cria no Dict [ dictRstArCtrls ], interno � classe [ clObjRstAreaParams ],
    ' um item referente a [ sCtrL ] pra guardar seu nome
    If clObjRstAreaParams.dictRstArCtrls.Exists(sCtrL) Then
        Set clObjLckdStatusParam = clObjRstAreaParams.dictRstArCtrls(sCtrL)
    
    Else
'Stop
        'Set clObjLckdStatusParam = New cls_07bLckdStatusParams
        clObjRstAreaParams.dictRstArCtrls.Add sCtrL, sCtrL

    End If
    '-----------------------------------------------------------------------------------------------------------------

'Stop


End Sub

Public Sub pbSub52_RstAreaBTNsDictBuild(vTagSection As Variant, cCtrL As Control)
    Dim vA, vB
    Dim vTagSectionParams As Variant
    Dim sParam As String
    Dim sStR1 As String, sStR2 As String
    Dim sForM As String
    Dim sCtrL As String
    Dim sBTNresetArea As String
    Dim sLoadLogWarn As String
    
    
    sCtrL = cCtrL.Name
    sForM = cCtrL.Parent.Name
    
'MsgBox "----- pbSub52_RstAreaBTNsDictBuild ---------------------------------------------" & vbCr & vbCr & "Confirma se o bot�o [ " & sCtrL & " ] est� configurado" & vbCr & "como [ Reset ]"
If gBbDepurandoLv01c Then Stop
'Stop
    
    
    vTagSectionParams = Split(vTagSection, ".")
    
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    sParam = "RstArea"

'parei aqui: checar por que a mensagem de alerta foi comentada

        'Mensagem de erro a ser inclu�da no Log de carga
        sLoadLogWarn = "Nenhuma �rea de reset foi atribu�da aos Bot�es a seguir. " & vbCrLf & "Eles n�o ir�o funcionar pra limpeza de campos no Formul�rio."
        
        'Mensagem de erro a ser exibida em tela
        'sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle:     " & " [ " & sCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
        'sStR2 = "Nenhuma �rea de reset foi atribu�da aos Bot�es a seguir." & vbCr & " Eles n�o ir�o funcionar pra limpar campos" & vbCr & " no Formul�rio:"

'MsgBox sStR1 & sStR2
'Stop
        On Error GoTo -1
        sBTNresetArea = GetTagParams(sParam, vTagSectionParams, , False, "", 0, , , , , True, "NoResetArForBtn", cCtrL, sLoadLogWarn)
        'sBTNresetArea = GetTagParams(sParam, vTagSectionParams, , "", , , sStR1, sStR2)
'Stop
        'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
        If sBTNresetArea = "" Then Exit Sub
        
'Stop
    '-------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------
    'Cria uma nova varia��o do dict [ dictRstArBTNsByNr(sForM ] pro Formul�rio corrente, caso ele ainda n�o tenha sido criado,
    ' pra guardar os bot�es de Reset do Form
    '-------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------
    If Not IsObject(dictRstArBTNsByNr(sForM)) Then Set dictRstArBTNsByNr(sForM) = New Dictionary
'Stop
    
    'Adiciona [ sBTNresetArea ] em [ dictRstArBTNsByNr(sForM) ]
    ' mas apenas se a �rea de Reset ainda n�o tiver sido inclu�da,
    ' do contr�rio da erro de duplicidade
    If Not dictRstArBTNsByNr(sForM).Exists(sBTNresetArea) = True Then
        
        'Antes de incluir o bot�o ao dict [ dictRstArBTNsByNr(sForM) ] confirma se existem Controles
        ' no formul�rio associados � �rea de Reset ora avalidada
        'Antes confirma se o dicicon�rio de controles em Areas de Reset existe
        If IsObject(dictFrmResetAreas(sForM)) Then
            
            'Confirma se a �rea de reset ora avaliada existe no dict [ dictFrmResetAreas(sForM) ]
            ' caso negativo significa que n�o h� controles associados a essa Reset Area
            ' e o bot�o n�o ser� inclu�do no dict [ dictRstArBTNsByNr(sForM  ]
            If dictFrmResetAreas(sForM).Exists(sBTNresetArea) = False Then

'parei aqui: checar por que a mensagem de alerta foi comentada

'Stop
'                vA = clObjTargtCtrlParam.sTargtCtrlName
'                sStR1 = "Formul�rio:      [ " & sForM & " ]" & vbCr & "�rea de Reset: " & " [ " & sBTNresetArea & " ]" & vbCr & "-------------------------------------------------------------------------------"
'                sStR2 = "N�o existem controles associados � essa �rea de Reset." & vbCr & " Esse bot�o n�o ir� esvaziar nenhum controle."
'                vA = " Erro [ " & Err.Number & ": " & Err.Description & " ] "
'
'                Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
                'Inclui o erro no dict de Logs de Carga do sistema
                sLoadLogWarn = "N�o h� controles associados � �rea de reset [ " & sBTNresetArea & " ], indicada nos bot�es a seguir." & vbCrLf & "Esses bot�es n�o ir�o esvaziar nenhum controle:"
                Call FormStatusBar01_Bld(sForM, "NoCtrlsInRstAr" & "_" & sBTNresetArea, sLoadLogWarn, sCtrL) 'NoCtrlsInRstAr
            
                Exit Sub
            
            End If
        
        
        End If
'Stop
        
       
       'Inclui o bot�o de Reset no dict [ dictRstArBTNsByNr(sForM) ]
        dictRstArBTNsByNr(sForM).Add sBTNresetArea, sCtrL
        
        'Inclui a mesma informa��o no dict [ dictRstArBTNsByName(sForM) ]
        If Not IsObject(dictRstArBTNsByName(sForM)) Then Set dictRstArBTNsByName(sForM) = New Dictionary
        If dictRstArBTNsByName(sForM).Exists(sCtrL) = False Then
            dictRstArBTNsByName(sForM).Add sCtrL, sBTNresetArea
        
        End If
        
    Else
        'A [ Reset Area ] ora avaliada j� existe no dict [ dictRstArBTNsByNr(sForM) ]

        'sStR1 = "Formul�rio:      [ " & sForM & " ]" & vbCr & "�rea de Reset: " & " [ " & sBTNresetArea & " ]" & vbCr & "-------------------------------------------------------------------------------"
        'sStR2 = "Bot�o associado em duplicidade � �rea de Reset." & vbCr & " Esse bot�o n�o ir� funcionar pra esvaziar" & vbCr & " nenhum controle."
        'vA = "  Bot�o de �rea de Reset duplicado"
        

    
        'Call msgboxErrorAlert(sStR1, sStR2, vbExclamation, vA)
'Stop
        'Inclui o erro no dict de Logs de Carga do sistema
        sLoadLogWarn = "Os bot�es a seguir foram associados em duplicidade � �rea de Reset [ " & sBTNresetArea & " ]." & vbCrLf & "Esses bot�es n�o ir�o esvaziar nenhum controle."
        Call FormStatusBar01_Bld(sForM, "DupBtnsForRstAr" & "_" & sBTNresetArea, sLoadLogWarn, sCtrL) 'NoCtrlsInRstAr
    
    End If
    
'Stop


End Sub


Public Sub pbSub60_CtrlsEnblDsblDictStartUp(fForM As Form)
    
    Dim vA, vB, vC
    Dim cEnblDsblCtrl As Control
    Dim sForM As String
    Dim sCtrL As String
    Dim sCtrlTAG As String
    Dim vSplittedTAG As Variant
    Dim iTagSection As Integer
    
'MsgBox "teste - carrega EnableDisable"
'Stop

    sForM = fForM.Name
    
    '-------------------------------------------------------------------------------------------------------
    'Cria uma nova varia��o do dicion�rio pro Formul�rio corrente, caso ele ainda n�o tenha sido criado
    ' pra depois carregar os par�metros originais dos Controles e tamb�m a vers�o desabilitada deles
    ' para serem aplicados durante a execu��o
    '-------------------------------------------------------------------------------------------------------
    If Not IsObject(dictCtrlEnblDsblParams(sForM)) Then Set dictCtrlEnblDsblParams(sForM) = New Dictionary
    
    
    'Loop pra localizar os Controles do [ Form ] sujeitos a "altera��o de comportamento" e manipular caracter�sticas como
    ' - EnableStatus
    ' - Highlight color
    ' - OnDirty color
    ' - TipText
    ' - DefaultColor
'Stop
    For Each cEnblDsblCtrl In fForM.Controls
        
        sCtrL = cEnblDsblCtrl.Name
    
'MsgBox "----- pbSub60_CtrlsEnblDsblDictStartUp -----------------------------------------" & vbCr & vbCr & "[ " & sCtrl & " ] est� na categoria de [ CtrlEnblDsbl ]?" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
'Stop
    
        Select Case cEnblDsblCtrl.ControlType
        
            'Avalia apenas os tipos de controles que podem apresentar altera��es de comportamento, que est�o sujeitos
            ' a diferentes n�veis de permiss�o ou podem ter o status Enabled alterado durante o funcionamento do sistema
            Case acCheckBox, acOptionButton, acTextBox, acListBox, acComboBox, acCommandButton, acToggleButton
            
                'Confirma se [ sCtrL ] tem a TAG necess�ria pra ser inclu�do em [ dictRstArBTNsByNr ] e em [ dictFormCommButtons(sForM) ]
                ' Chama a rotina pra montar ambos dicion�rios
                If cEnblDsblCtrl.ControlType = acCommandButton Then

'Stop
                    sCtrlTAG = cEnblDsblCtrl.Tag
                    
vA = "----- pbSub60_CtrlsEnblDsblDictStartUp -----------------------------------------" & vbCr & vbCr & "Avalia se [ " & sCtrL & " ] tem a TAG necess�ria pra "
vB = vbCr & "inclus�o no dict [ dictRstArBTNsByNr(sForm) ]" & vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
'MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop

                    'Se o Controle for mesmo um TriggerCtrl chama rotina pra guardar os par�metros no Dict.
                    ' Devem ser armazenados diversos par�metros necess�rios pra alterar o SQL do controle
                    ' e fazer a filtragem
                    If InStr(1, sCtrlTAG, "RstArea") > 0 Then
                    
                    'Chama rotina pra montar o dicion�rio [ dictRstArBTNsByNr(sForm) ]
                    '-------------------------------------------------------------------------------------------------------------
                    '----------------------------------------------
                        'Separa os par�metros do controle (no caso do ResetBtn apenas um par�metro)
                        vSplittedTAG = Split(sCtrlTAG, "-")
                        iTagSection = 1

                        'Avalia a 1a se��o com par�metros de TrggCtrl
                        If vSplittedTAG(iTagSection - 1) <> "" Then
'Stop
                            Call pbSub52_RstAreaBTNsDictBuild(vSplittedTAG(iTagSection - 1), cEnblDsblCtrl)
                            
                        End If

                    End If
                    
vA = "----- pbSub60_CtrlsEnblDsblDictStartUp -----------------------------------------" & vbCr & vbCr & "Avalia se [ " & sCtrL & " ] tem a TAG necess�ria pra "
vB = vbCr & "inclus�o no dict [ dictFormCommButtons(sForm) ]" & vbCr & vbCr & "TAG [" & Chr(160) & sCtrlTAG & Chr(160) & "]"
If gBbDepurandoLv01b Then MsgBox vA & vB
If gBbDepurandoLv01b Then Stop
'Stop
                    

                    'Pra todos os bot�es do [ form ] chama rotina pra montar o dicion�rio [ dictFormCommButtons(sForM) ]
                    ' e, se atendidos os requisitos, incluir os bot�es no dict [ dictCtrlsEvents(sForM) ]
                    '-------------------------------------------------------------------------------------------------------------
                    '----------------------------------------------
                    Call pbSub81_CommButtonsEventBuild(sCtrlTAG, cEnblDsblCtrl)
                
                Else
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Dict Events build"
'Stop
                    '-----------------------------------------------------
                    'Chama a rotina pra incluir o controle no dict [ dictCtrlsEvents(sForM) ]
                    Call pbSub10_EventsDictBuild(sForM, sCtrL)
                    '-----------------------------------------------------
                
                End If
        
                
                'sCtrlTAG = cEnblDsblCtrl.Tag
'Stop
    
                'Chama a rotina pra guardar os par�metros no Dict.
                ' Guarda por exemplo, o tipText original do controle no caso dele ser alterado conforme situa��es no sistema
                    If gBbDebugOn Then Debug.Print " " & sCtrL
'Stop
                    
If gBbDepurandoLv01b Then MsgBox "----- pbSub60_CtrlsEnblDsblDictStartUp -----------------------------------------" & vbCr & vbCr & "Chama [ pbSub61_cCtrlsEnblDsblDictBuild ] pra incluir" & vbCr & "[ " & sCtrL & " ] no dict [ dictCtrlEnblDsblParams(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01b Then Stop
                    
                    On Error GoTo -1
                    Call pbSub61_cCtrlsEnblDsblDictBuild(sForM, sCtrlTAG, cEnblDsblCtrl)
'Stop
        Case acLabel
        
            'Somente a label "lblStatusBar" de cada formul�rio � inclu�da no dict [ dictCtrlsEvents(sForM) ] pra
            ' que o duplo clique abra o log de carga do sistema
            If cEnblDsblCtrl.Name = bBsStatusBarLabel Then
                '-----------------------------------------------------
                'Chama a rotina pra incluir o controle no dict [ dictCtrlsEvents(sForM) ]
                Call pbSub10_EventsDictBuild(sForM, sCtrL)
                '-----------------------------------------------------
            
            End If
            
        End Select
    
    Next cEnblDsblCtrl
'Stop

End Sub

Public Sub pbSub61_cCtrlsEnblDsblDictBuild(sForM As String, sCtrlTAG As String, cEnblDsblCtrl As Control)
         
    Dim vA, vB, vC
    Dim sCtrL As String
    'Dim sForM As String
    Dim dDicT As Dictionary
    Dim sCtrlType As String
    Dim sLockdStatus As String
    
    '---------------------------------------------------------------------------------------
    'Armazena no dict  [ ] os par�metros pra Habilitar e Desabilitar os controles do sistema
    '---------------------------------------------------------------------------------------
    
    sCtrL = cEnblDsblCtrl.Name
    sCtrlType = dictCtrlTypeShort(cEnblDsblCtrl.ControlType)
    
If gBbDepurandoLv01c Then MsgBox "----- pbSub61_cCtrlsEnblDsblDictBuild ------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sCtrL & " ] pra inclus�o" & vbCr & "em [ dictCtrlEnblDsblParams(sForM) ]"
If gBbDepurandoLv01c Then Stop
    
'If gBbDepurandoLv01b Then MsgBox "teste - Get Enbl/Dsbl params [ " & sCtrL & " ]"
If gBbDepurandoLv01b Then Stop
        'A rotina que chamou esse procedimento j� confirmou a exist�ncia do dicion�rio
        ' portanto pode-se usar Set dDict direto
        Set dDicT = dictCtrlEnblDsblParams(sForM)
        
        'Na eventualidade do controle j� ter sido inclu�do no dicion�rio Dict [ dictCtrlEnblDsblParams ]
        ' apenas recupera pra guardar os par�metros
        If dDicT.Exists(sCtrL) = True Then
                Set clObjCtrlsEnblDsblParams = dDicT(sCtrL)
Stop
        Else
            'O Controle ainda n�o foi inclu�do no Dict ent�o cria um novo
            ' objeto [ clObjCtrlsEnblDsblParams ] da Classe [ cls_07aCtrlsEnblDsblParams ] pra guardar os par�metros
            ' e em seguida adicona o Controle ao Dict
            Set clObjCtrlsEnblDsblParams = New cls_07aCtrlsEnblDsblParams
        
        End If
            
        'Adiciona o Controle no dicion�rio [ dictCtrlEnblDsblParams ] e guarda nele o objeto [ clObjCtrlsEnblDsblParams ]
        ' com os respectivos par�metros de Habilita��o/Desabilita��o definidos na classe [ cls_07aCtrlsEnblDsblParams ]
        dDicT.Add sCtrL, clObjCtrlsEnblDsblParams

'If gBbDepurandoLv01b Then MsgBox "teste - Guardar os par�metros comuns"
'stop
        '-----------------------------------------------------------------------------------------------------------------
        '-------------------------------------------
        'Armazena os par�metros do Controle esperados pela Classe [ cls_07aCtrlsEnblDsblParams ]
        ' par�metros comuns aos dois status: Enabled/Disabled
        With clObjCtrlsEnblDsblParams

'Stop
            .sCtrlName = cEnblDsblCtrl.Name
            .sInitTipText = cEnblDsblCtrl.ControlTipText
            
            If sCtrlType = "btn" Then 'Controle tipo CommandButton
                'Apenas controles CommandButton aceitam a propriedade Gradient
                .iCtrlGradientColor = cEnblDsblCtrl.Gradient '25 como padr�o
            
            End If

        End With
        
        '-------------------------------------------
        '-----------------------------------------------------------------------------------------------------------------
        If dictCtrlEnblDsblParams(sForM)(sCtrL).dictParamByLckdStatus Is Nothing Then Set dictCtrlEnblDsblParams(sForM)(sCtrL).dictParamByLckdStatus = New Dictionary
        
        '-----------------------------------------------------------------------------------------------------------------
        '-------------------------------------------
        'Cria no Dict [ dictParamByLckdStatus ], dentro da classe [ clObjCtrlsEnblDsblParams ],
        ' um item [ "Enbld" ] e p�e dentro dele o objeto [ dictParamByLckdStatus ] com os par�metros
        ' da classe [ cls_07bLckdStatusParams ] para a vers�o [ Enabled ] do controle
        sLockdStatus = "Enbld"
        If clObjCtrlsEnblDsblParams.dictParamByLckdStatus.Exists(sLockdStatus) Then
            Set clObjLckdStatusParam = clObjCtrlsEnblDsblParams.dictParamByLckdStatus(sLockdStatus)
        
        
        Else
'Stop
            Set clObjLckdStatusParam = New cls_07bLckdStatusParams
            clObjCtrlsEnblDsblParams.dictParamByLckdStatus.Add sLockdStatus, clObjLckdStatusParam
            
        End If
        
        'Cria no Dict [ dictParamByLckdStatus ], dentro da classe [ clObjCtrlsEnblDsblParams ],
        ' um item [ "Dsbld" ] e p�e dentro dele o objeto [ dictParamByLckdStatus ] com os par�metros
        ' da classe [ cls_07bLckdStatusParams ] para a vers�o [ Disabled ] do controle
        sLockdStatus = "Dsbld"
        If clObjCtrlsEnblDsblParams.dictParamByLckdStatus.Exists(sLockdStatus) Then
            Set clObjLckdStatusParam = clObjCtrlsEnblDsblParams.dictParamByLckdStatus(sLockdStatus)
        
        Else
'Stop
            Set clObjLckdStatusParam = New cls_07bLckdStatusParams
            clObjCtrlsEnblDsblParams.dictParamByLckdStatus.Add sLockdStatus, clObjLckdStatusParam
            
        End If

'If gBbDepurandoLv01b Then MsgBox "teste - Guardar os par�metros [ Enbld ]"
'Stop
        '-----------------------------------------------------------------------------------------------------------------
        '-------------------------------------------
        'Armazena os par�metros do Controle esperados pela Classe [ cls_07bLckdStatusParams ] para a vers�o [ Enabled ]
        Set clObjLckdStatusParam = clObjCtrlsEnblDsblParams.dictParamByLckdStatus("Enbld")
        With clObjLckdStatusParam
        
'Stop
            If sCtrlType <> "chk" And sCtrlType <> "opb" And sCtrlType <> "acOptionGroup" Then 'controle n�o � Checkbox nem Optionbutton
                .lngLckdStatusBackColor = cEnblDsblCtrl.BackColor
                .lngLckdStatusForeColor = cEnblDsblCtrl.ForeColor
                .lngLckdStatusBorderColor = cEnblDsblCtrl.BorderColor    'necess�rio apenas pra Commandbutton
                .lngLckdStatusBorderStyle = cEnblDsblCtrl.BorderStyle
                
            End If
            
            If sCtrlType <> "btn" And sCtrlType <> "" Then 'qualquer controle exceto CommandButton
                .iLckdStatusSpecialEffect = cEnblDsblCtrl.SpecialEffect
            
            End If
            
            'Debug.Print "BackColor: " & .lngLckdStatusBackColor
            'Debug.Print "ForeColor: " & .lngLckdStatusForeColor
            'Debug.Print "BorderColor: " & .lngLckdStatusBorderColor
            'Debug.Print "SpecialEffect: " & .iLckdStatusSpecialEffect
            'Debug.Print "BorderStyle: " & .lngLckdStatusBorderStyle
            
        End With
        
        
'If gBbDepurandoLv01b Then MsgBox "teste - Guardar os par�metros [ Dsbld ]"
'stop
        '-----------------------------------------------------------------------------------------------------------------
        '-------------------------------------------
        'Armazena os par�metros do Controle esperados pela Classe [ cls_07bLckdStatusParams ] para a vers�o [ Disabled ]
        Set clObjLckdStatusParam = clObjCtrlsEnblDsblParams.dictParamByLckdStatus("Dsbld")
        With clObjLckdStatusParam
'Stop
            'Par�metros da vers�o [ Enabled ]
            If sCtrlType <> "chk" And sCtrlType <> "opb" Then 'controle n�o � Checkbox nem Optionbutton
            
                If sCtrlType = "btn" Then 'controle tipo CommandButton
                    'Cores diferentes pra Commandbutton e pra Txt, List e Combobox
                    .lngLckdStatusBackColor = GbLngBtnGREyBackColor
                    .lngLckdStatusBorderColor = GbLngBtnGREyBdColor
                    .lngLckdStatusBorderStyle = cEnblDsblCtrl.BorderStyle
                
                Else  'controle tipo Text, List ou Combobox
                    .lngLckdStatusBackColor = GbLngTxtBaseBackColor  'n�o necess�rio pra Checkbox e OptionButton
                    .lngLckdStatusForeColor = GbLngTxtBASeForeColor  'n�o necess�rio pra Checkbox e OptionButton
                
                    .iLckdStatusSpecialEffect = 0
                
                End If
            
            Else         'controle tipo Checkbox ou Optionbutton
                .iLckdStatusSpecialEffect = 5
                .lngLckdStatusBorderStyle = cEnblDsblCtrl.BorderStyle
            
            End If
'Stop
            'Debug.Print "BackColor: " & .lngLckdStatusBackColor
            'Debug.Print "ForeColor: " & .lngLckdStatusForeColor
            'Debug.Print "BorderColor: " & .lngLckdStatusBorderColor
            'Debug.Print "SpecialEffect: " & .iLckdStatusSpecialEffect
            'Debug.Print "BorderStyle: " & .lngLckdStatusBorderStyle
        
        
        End With
        '-------------------------------------------
        '-----------------------------------------------------------------------------------------------------------------

'Stop
End Sub


Public Sub pbSub71_DataFieldDictBuild(vTagSection As Variant, cDataField As Control)
    Dim vA, vB
    Dim sDataField As String
    Dim sForM As String
    Dim sCtrlDataField As String
    Dim sDataFieldGrp As String
    Dim sRecQry As String
    Dim sStR1 As String, sStR2 As String, sLoadLogWarn As String
    Dim sParam As String
    Dim vTagSectionParams As Variant
    Dim iInT As Integer

    sForM = cDataField.Parent.Name
    sCtrlDataField = cDataField.Name
    vA = Split(vTagSection, ".")

    '-------------------------------------------------------------------------------------------------------
    'Recupera os par�metros do controle informados na TAG
    '-------------------------------------------------------------------------------------------------------

    sParam = "DataField>"
    For iInT = 0 To UBound(vA)
        If InStr(vA(iInT), sParam) > 0 Then vTagSectionParams = vA(iInT)
    Next iInT
    If vTagSectionParams <> "" Then sDataField = Split(vTagSectionParams, ">")(1)
    
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do TrggCtrl
    sParam = "Grp"
    'Mensagem de erro a ser inclu�da no Log de carga
    sLoadLogWarn = "Os controles DataField a seguir n�o est�o associados a nenhum grupo e n�o poder�o ser manipulados."
    
    'Mensagem de erro a ser exibida em tela
    sStR1 = "Formul�rio:  [ " & sForM & " ]" & vbCr & "Controle: " & "   [ " & sCtrlDataField & " ]" & vbCr & "-------------------------------------------------------------------------------"
    sStR2 = " O [ Grupo ] do [ CtrlDataField ] n�o foi informado" & vbCr & "  e ele n�o poder� ser manipulado."

    sDataFieldGrp = GetTagParams(sParam, vA, , True, "", 1, , True, sStR1, sStR2, True, "DataFieldGrpError", cDataField, sLoadLogWarn)

    sParam = "RecQry>"
    For iInT = 0 To UBound(vA)
        If InStr(vA(iInT), sParam) > 0 Then vTagSectionParams = vA(iInT)
    Next iInT
    sRecQry = GetTagParams(sParam, vA, , , "")

    If sDataFieldGrp <> "" Then
    
        If Not IsObject(dictFormDataFlds00Ctrls(sForM)) Then Set dictFormDataFlds00Ctrls(sForM) = New Dictionary
        If Not dictFormDataFlds00Ctrls.Exists(sForM) Then Stop
        
        'Adiciona [ sCtrlDataField ] em [ dictFormDataFlds00Ctrls(sForM) ]
        If dictFormDataFlds00Ctrls(sForM).Exists(sCtrlDataField) = True Then
            Set clObjCtrlDataFieds = dictFormDataFlds00Ctrls(sForM)(sCtrlDataField)
    
        Else
           'Cria um novo objeto [ clObjCtrlDataFieds ] da Classe [ cls_04aCtrlsDataFields ] pra ser inclu�do no Dict
            Set clObjCtrlDataFieds = New cls_04aCtrlsDataFields
            
            dictFormDataFlds00Ctrls(sForM).Add sCtrlDataField, clObjCtrlDataFieds
            clObjCtrlDataFieds.sCtrlDataField = sCtrlDataField
            clObjCtrlDataFieds.sDataFieldGrp = sDataFieldGrp
            clObjCtrlDataFieds.sDataField = sDataField
            clObjCtrlDataFieds.sRecQry = sRecQry
            
        End If
        
        If Not IsObject(dictFormDataFlds01Grps(sForM)) Then Set dictFormDataFlds01Grps(sForM) = New Dictionary
        If Not IsObject(dictFormDataFlds01Grps(sForM)(sDataFieldGrp)) Then Set dictFormDataFlds01Grps(sForM)(sDataFieldGrp) = New Dictionary
        
        dictFormDataFlds01Grps(sForM)(sDataFieldGrp).Add sCtrlDataField, clObjCtrlDataFieds
        
        If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)(sDataFieldGrp)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM)(sDataFieldGrp) = New Dictionary
        If Not dictFormFilGrpsEnDsAllCtrls(sForM)(sDataFieldGrp).Exists(sCtrlDataField) Then dictFormFilGrpsEnDsAllCtrls(sForM)(sDataFieldGrp).Add sCtrlDataField, sDataFieldGrp
    End If

End Sub


Public Sub pbSub81_CommButtonsEventBuild(vTagSection As Variant, cCommButton As Control)
    Dim vA, vB
    
    Dim vTagSectionParams As Variant
    Dim sCommButton As String
    Dim sForM As String
    Dim sStR1 As String, sStR2 As String
    Dim sLoadLogWarn As String
    Dim sParam  As String
    Dim sFilGrp As String
    Dim sActType As String
    Dim sFrmMode As String
    Dim sFormToOpen As String
    Dim sRecQry As String
    Dim sRstArea As String
    Dim sColToSort As String
    Dim sOrderMode As String
    
    sCommButton = cCommButton.Name
    sForM = cCommButton.Parent.Name
    
    '-------------------------------------------------------------------------------------------------------
    'Recupera os par�metros do controle informados na TAG
    '-------------------------------------------------------------------------------------------------------
                                 
'MsgBox "----- pbSub81_CommButtonsEventBuild ---------------------------------------------" & vbCr & vbCr & "Recupera os par�metros de [ " & sCommButton & " ]" & vbCr & "do form [ " & sForm & " ] pra inclus�o" & vbCr & "em [ dictFormCommButtons(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01c Then Stop
'Stop
    
    If vTagSection = "" Then
        'Mensagem de erro a ser inclu�da no Log de carga
        vA = "Nos seguintes bot�es n�o foram localizados nenhum par�metro de funcionamento." & vbCrLf
        vB = "Esses bot�es n�o ir�o funcionar."
        sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
        On Error GoTo -1
        Call FormStatusBar01_Bld(sForM, "commBtn-NoParams", sLoadLogWarn, sCommButton)
        'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
        
        Exit Sub
    
    'O termo [ NotInDict ] inclu�do na TAG de bot�es indica que eles
    ' n�o devem ser inclu�dos no dict de bot�es [ dictFormCommButtons(sForM) ] nem no dict de eventos [ dictCtrlsEvents(sForM) ]
    ElseIf vTagSection = "NotInDict" Then Exit Sub
    
    End If
    
    
    vTagSectionParams = Split(vTagSection, ".")
    
    '-----------------------------------------------------------------------------------------------------
    ' Recupera os par�metros do [ cCommButton ]
    '-----------------------------------------------------------------------------------------------------
    'Verifica se foi identificado o par�metro [ RstArea ] contendo o [ Grupo de Filtragem ] do cCommButton
    sParam = "RstArea"
        
        On Error GoTo -1
        sRstArea = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
'Stop
    'Verifica se foi identificado o par�metro [ sFilGrp ] contendo o [ Grupo de Filtragem ] do cCommButton
    sParam = "Grp"
        
        On Error GoTo -1
        sFilGrp = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
        
'Stop
    
    'Verifica se foi identificado o par�metro [ ActType ] contendo o [ Grupo de Filtragem ] do cCommButton
    sParam = "ActType>"

        On Error GoTo -1
        sActType = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
'Stop
    
    'Verifica se foi identificado o par�metro [ FrmMode ] contendo o [ Grupo de Filtragem ] do cCommButton
    sParam = "FormMode>"
    
        On Error GoTo -1
        sFrmMode = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
    
    'Verifica se foi identificado o par�metro [ RecQry> ] contendo a [ Consulta ] que ser� utilizada pelo cCommButton
    sParam = "RecQry>"

        On Error GoTo -1
        sRecQry = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
        Debug.Print "pbsub81 - " & cCommButton.Name
'Stop
    
    'Verifica s'e foi identificado o par�metro [ Form ] contendo o [ Grupo de Filtragem ] do cCommButton
    sParam = "Form>"
    
        On Error GoTo -1
        sFormToOpen = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
        'sFormToOpen = GetTagParams(sParam, vTagSectionParams, , True, "", 1, , True, sStR1, sStR2, True, "MissingTrggFilGrp", cCommButton, sLoadLogWarn)
'Stop
    sParam = "OrderMode>"
        sOrderMode = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
    
    sParam = "ColToSort>"
        sColToSort = GetTagParams(sParam, vTagSectionParams, , False, "", 1, , False, , , False, , cCommButton)
        
    '-----------------------------------------------------------------------------------------------------
    'Ap�s recuperar os par�metros da [ Tag ] do [ cCommButton ]
    ' verifica se os requisitos foram cumpridos, dependendo do tipo de a��o atribu�da ao bot�o
    ' caso negativo adiciona o alerta no log de carga do sistema
    '-----------------------------------------------------------------------------------------------------
    Select Case sActType
        
        'Os seguintes tipos de bot�o exigem a indica��o do [ Grp de Filtragem ]
        Case "EditRec", "NewRec", "SaveEdit", "SaveNew", "CancelEdit"
'Stop
            'Caso seja um bot�o de salvamento de dados dever� obrigat�riamente ter a consulta para salvamento indicada
            '  na TAG pelo par�metro [ RecQry> ], caso o par�metro n�o seja encontrado, carrega o erro no log de carga
            '  e exibe mensagem em tela
            If InStr(sActType, "Save") > 0 Then

                If sRecQry = "" Then
Stop
                    vA = "Os seguintes bot�es est�o atribu�dos � a��o de" & vbCrLf & "[" & Chr(160) & "Salvamento de Registros" & Chr(160) & "] "
                    vB = "mas n�o possuem o par�metro [ RecQry> ] " & vbCrLf & "impossibilitando a grava��o de dados."
                    sLoadLogWarn = vA & vB

                    On Error GoTo -1
                    Call FormStatusBar01_Bld(sForM, "commBtn-MissingRecQry", sLoadLogWarn, sCommButton)
                    
                    Exit Sub
                    
                End If
            End If
            If sFilGrp = "" Then
                vA = "Os seguintes bot�es est�o atribu�dos � a��o de" & vbCrLf & "[" & Chr(160) & "Salvamento de Registros" & Chr(160) & "] "
                vB = "mas n�o est�o associados a nenhum" & vbCrLf & "grupo de filtragem e por isso n�o ir�o funcionar."
                sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                On Error GoTo -1
                Call FormStatusBar01_Bld(sForM, "commBtn-MissingFilGrp", sLoadLogWarn, sCommButton)
                'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                
                Exit Sub
            
            End If
        
        'Os seguintes tipos de bot�o exigem a indica��o do [ Form ]
        Case "OpenForm"
'Stop
            If sFormToOpen = "" Then
'Stop
                'Mensagem de erro a ser inclu�da no Log de carga
                vA = "Os seguintes bot�es est�o atribu�dos � a��o de" & vbCrLf & "[" & Chr(160) & "Abertura de Formu�rio" & Chr(160) & "] "
                vB = "mas o formul�rio" & vbCrLf & "respectivo n�o foi informado e por isso n�o ir�o funcionar."
                sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                On Error GoTo -1
                Call FormStatusBar01_Bld(sForM, "commBtn-MissingForm", sLoadLogWarn, sCommButton)
                'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                
                Exit Sub
                
            End If
            
            'Se for tipo [ OpenForm ] exige tamb�m o [ FrmMode ]
            If sActType = "OpenForm" Then
'Stop
                If sFrmMode = "" Then
'Stop
                    'Mensagem de erro a ser inclu�da no Log de carga
                    vA = "Os seguintes bot�es est�o atribu�dos � a��o de" & vbCrLf & "[" & Chr(160) & "Abertura de Formu�rio" & Chr(160) & "] "
                    vB = "mas o formul�rio respectivo" & vbCrLf & "n�o foi informado e por isso n�o ir�o funcionar."
                    sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                    On Error GoTo -1
                    Call FormStatusBar01_Bld(sForM, "commBtn-MissingFrmMode", sLoadLogWarn, sCommButton)
                    'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                    
                    Exit Sub

                End If
                
            End If
        
        'O seguinte tipo de bot�o exige a indica��o da [ ResetArea ]
        Case "RstArea"
Stop
            If sRstArea = "" Then
Stop
                'Mensagem de erro a ser inclu�da no Log de carga
                vA = "Os seguintes bot�es est�o atribu�dos � a��o de" & vbCrLf & "[" & Chr(160) & "Reset Area" & Chr(160) & "] "
                vB = "mas n�o est�o associados a nenhuma" & vbCrLf & "[ Reset Area ] e por isso n�o ir�o funcionar."
                sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
                On Error GoTo -1
                Call FormStatusBar01_Bld(sForM, "commBtn-MissingRstArea", sLoadLogWarn, sCommButton)
                'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
                
                Exit Sub
            
            End If
            
        Case ""
                    
            'Verifica se foi identificado um valor v�lido pro par�metro [ sActType ]
            ' do contr�rio carrega o erro no log e n�o inclui o [ bot�o ] no dict
'Stop
            'Mensagem de erro a ser inclu�da no Log de carga
            vA = "Nos seguintes bot�es n�o foi identificado um valor v�lido pro par�metro [" & Chr(160) & "sActType" & Chr(160) & "] " & vbCrLf
            vB = "Esses bot�es n�o ir�o funcionar."
            sLoadLogWarn = vA & vB
'MsgBox sLoadLogWarn
'Stop
            On Error GoTo -1
            Call FormStatusBar01_Bld(sForM, "commBtn-MissingActType", sLoadLogWarn, sCommButton)
            'If gBbEnableErrorHandler Then On Error GoTo -1: On Error GoTo FrM_ErrorHandler
            
            Exit Sub
            
    End Select
'Stop
'parei aqui:2
    '-----------------------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------------------
    ' Inclui [ cCommButton ] no dict [ dictFormCommButtons(sForM) ]
    '-----------------------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------------------
    If Not IsObject(dictFormCommButtons(sForM)) Then Set dictFormCommButtons(sForM) = New Dictionary
    
    If dictFormCommButtons(sForM).Exists(sCommButton) Then
        Set clObjCommButtons = dictFormCommButtons(sForM)(sCommButton)
    
    Else
        Set clObjCommButtons = New cls_12aCommButtonParams
            dictFormCommButtons(sForM).Add sCommButton, clObjCommButtons
                
    End If

'Stop
    '
    clObjCommButtons.sCtrlName = sCommButton
    clObjCommButtons.sActType = sActType
    clObjCommButtons.sFilGrp = sFilGrp
    clObjCommButtons.sRstArea = sRstArea
    clObjCommButtons.sForM = sFormToOpen
    clObjCommButtons.sFrmMode = sFrmMode
    clObjCommButtons.sRecQry = sRecQry
    clObjCommButtons.sColToSort = sColToSort
    clObjCommButtons.sOrderMode = sOrderMode
    
    '-----------------------------------------------------
    'Chama a rotina pra incluir o controle no dict [ dictCtrlsEvents(sForM) ]
    Call pbSub10_EventsDictBuild(sForM, sCommButton)
    '-----------------------------------------------------
    If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM) = New Dictionary
    If Not IsObject(dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp)) Then Set dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp) = New Dictionary
    If Not dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Exists(sCommButton) Then dictFormFilGrpsEnDsAllCtrls(sForM)(sFilGrp).Add sCommButton, sFilGrp
    
End Sub
