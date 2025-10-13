Attribute VB_Name = "M�dulo 08a - pbSubsEnblDsbl"
Option Compare Database
Option Explicit


Public Sub pbSub00_CtrlsEnblDsble_GetParams(fForM As Form, sSysFormMode As String, Optional sFilGrp As String, Optional cPressedCtrl As Control)
    Dim vA, vB, vC, vD, vE, vF
    Dim rsTbE As Recordset
    Dim sQuerY As String
    Dim sWhere As String
    Dim lngFoundRecs As Long
    Dim sSysForM As String
    Dim sPressedCtrL As String
    Dim bEnable As Boolean
    Dim bLockCombo As Boolean
    Dim cTweakableCtrL As Control
    Dim sTweakableCtrL As String
    Dim sCtrlGrp As String
    Dim bVisible As Boolean
    Dim sCtrlNewTipText As String
    Dim bPermissionGrated As Boolean
    Dim sLoadLogWarn As String
    Dim sForM As String
    
    
    sForM = fForM.Name
    
    'Recupera na tabela [ qry_00(3)bSysEnblDisblParms(Edt) ] os par�metros necess�rios
    ' com a indica��o dos ajustes a serem feitos nos Controles do [ Form ]
    ' a partir da a��o iniciada (ex.: FormLoad, pressionamento do bot�o Editar, etc)
    
    
    sSysForM = fForM.Name
    If Not cPressedCtrl Is Nothing Then sPressedCtrL = cPressedCtrl.Name
    
    
    'Indica a Consulta que tem os par�metros de Enable/Disable para serem recuperados
    sQuerY = "qry_01(03)bSysEnblDisblParams"
    
    'Monta a filtragem
    vA = "([sSysForM] Like " & """" & sSysForM & """" & ")"
     'Debug.Print vA
    vB = IIf(Not cPressedCtrl Is Nothing, " And " & "([sTriggerCtrl] Like " & """" & sPressedCtrL & """" & ")", "")
     'Debug.Print vB
    vC = IIf(Not IsNull(sSysFormMode), " And " & "([sSysFormMode] Like " & """" & sSysFormMode & """" & ")", "")
     'Debug.Print vC
    sWhere = vA & vB & vC
     If gBbDebugOn Then Debug.Print sWhere

'Stop
    
    'Abre a consulta e aplica o filtro [ sWhere ]
    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    rsTbE.Filter = sWhere
    Set rsTbE = rsTbE.OpenRecordset
    
    'Apenas para checar a quantidade de registros. Desnecess�rio nessa rotina
    lngFoundRecs = rsTbE.RecordCount
    If lngFoundRecs > 0 Then
        rsTbE.MoveLast
        lngFoundRecs = rsTbE.RecordCount
    End If
    
    'Passa por todos os controles indicados na tabela [ qry_00(3)bSysEnblDisblParms(Edt) ]
    ' pra identificar as altera��es necess�rias
    '
    ' Em seguida chama rotina pra aplicar a altera��o dos controles [ sTweakbleCtrl ]
    ' conforme par�metros definidos na tabela de dados [ sQuery ]
    If Not (rsTbE.EOF And rsTbE.BOF) Then
        rsTbE.MoveFirst
        Do Until rsTbE.EOF = True
'Stop
            
            sTweakableCtrL = rsTbE.Fields("sTweakbleCtrl")
            
            If ControlExists(sTweakableCtrL, Forms(sSysForM)) Then
                Set cTweakableCtrL = Forms(sSysForM).Controls(sTweakableCtrL)
'Stop
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Enbl/Dsble Get params: [ " & sTweakableCtrL & " ]"
'Stop
                
                bEnable = rsTbE.Fields("bEnable")
                bVisible = rsTbE.Fields("bVisible")
                bLockCombo = rsTbE.Fields("bLockCombo")
    
                vA = Not IsNull(rsTbE.Fields("sAltTipText"))
                If vA Then sCtrlNewTipText = rsTbE.Fields("sAltTipText")
    
                'Verifica se o usu�rio tem permiss�o para que o controle seja habilitado
                If bEnable Then
'MsgBox "teste - Check permissions before enable: [ " & sTweakableCtrL & " ]"
'Stop
                    GetPrmissGrntedType = bCheckUserPermissionLevel(sSysForM, sTweakableCtrL)
    
                    'Se o usu�rio n�o tiver a permiss�o requerida pra acessar o Controle
                    If Not GetPrmissGrntedType.bPermissionGrated Then
                        
                        bEnable = False
                        sCtrlNewTipText = GetPrmissGrntedType.sCtrlNewTipText
                        'cTweakableCtrL.ControlTipText = GetPrmissGrntedType.sCtrlNewTipText
                    
                    End If

'Stop
                    'vA = GetPrmissGrntedType.bPermissionGrated
                    'vB = GetPrmissGrntedType.sCtrlNewTipText
'Stop
                End If
                    
'                vA = Not IsNull(rsTbE.Fields("sAltTipText"))
'                If vA Then sCtrlNewTipText = rsTbE.Fields("sAltTipText")
                    
'                'Se o usu�rio n�o tiver a permiss�o requerida pra acessar o Controle
'                If Not GetPrmissGrntedType.bPermissionGrated Then
'
'                    bEnable = False
'                    sCtrlNewTipText = GetPrmissGrntedType.sCtrlNewTipText
'                    'cTweakableCtrL.ControlTipText = GetPrmissGrntedType.sCtrlNewTipText
'
'                End If
        
            
'Stop
                vA = fForM.Name
                vB = cTweakableCtrL.Name
                
                'Se o grupo de filtragem for indicado por [ sFilGrp ] na chamada da rotina
                ' verifica se o controle pertence ao grupo, caso positivo prossegue com o Enable/Disable do controle
                If sFilGrp <> "" Then
'Stop
                    If dictFormFilterGrpsCtrls(sForM)(sFilGrp).Exists(sTweakableCtrL) Then
                        Call pbSub01_CtrlsEnblDsble_Confirm(fForM, cTweakableCtrL, dictCtrlTypeShort(cTweakableCtrL.ControlType), bEnable, bVisible, bLockCombo, sCtrlNewTipText)
                    End If
                Else
                    Call pbSub01_CtrlsEnblDsble_Confirm(fForM, cTweakableCtrL, dictCtrlTypeShort(cTweakableCtrL.ControlType), bEnable, bVisible, bLockCombo, sCtrlNewTipText)
                End If
            Else
                'Carrega pro Log de carga do sistema os controles indicados na tabela que n�o existam no [ Form ]
                sLoadLogWarn = "Os seguintes controles foram indicados na tabela [" & Chr(160) & "qry_01(03)bSysEnblDisblParams" & Chr(160) & "] para configura��o de Status Enabled mas n�o existem no formul�rio"
                Call FormStatusBar01_Bld(sForM, "MissingEnblDsblCtrls", sLoadLogWarn, sTweakableCtrL)

'Stop

                        
            End If
'Stop
'MyNextRecord:
'            'Move to the next record
            rsTbE.MoveNext
        
        Loop
    
    End If
'Stop
    rsTbE.Close 'Close the recordset
    Set rsTbE = Nothing 'Clean up

End Sub

Public Sub pbSub01_CtrlsEnblDsble_Confirm(ByVal fForM As Form, cTweakableCtrL As Control, sCtrlType As String, bEnable As Boolean, bVisible As Boolean, Optional bLocked As Boolean, Optional sCtrlNewTipText As String)

    Dim vA, vB
    Dim cOptBttn As Control
    Dim sForM As String
    
    
    'Antes de entrar na Rotina [ pbSub02_CtrlsEnblDsbl_Apply ] pra fazer a altera��o dos controles
    ' desmembra os controls [ acOptionGroup ] pra aplicar as altera��es nos
    ' itens internos ao invez de aplicar no controle Pai
    
    sForM = fForM.Name
    vA = cTweakableCtrL.Name

'MsgBox "teste - aplicando Enable/Disable para: [ " & vA & " ]"
'Stop
    If sCtrlType = "grp" Then
        
        For Each cOptBttn In cTweakableCtrL.Controls
            vA = cOptBttn.Name
'Stop
            Call pbSub02_CtrlsEnblDsbl_Apply(fForM, cOptBttn, dictCtrlTypeShort(cOptBttn.ControlType), bEnable, bVisible, bLocked, sCtrlNewTipText)
            
        Next cOptBttn
    
    
    Else
        Call pbSub02_CtrlsEnblDsbl_Apply(fForM, cTweakableCtrL, dictCtrlTypeShort(cTweakableCtrL.ControlType), bEnable, bVisible, bLocked, sCtrlNewTipText)

    End If


End Sub


Public Sub pbSub02_CtrlsEnblDsbl_Apply(ByVal fForM As Form, cTweakableCtrL As Control, sCtrlType As String, bEnable As Variant, bVisible As Boolean, Optional bLocked As Boolean, Optional sCtrlNewTipText As String)

    Dim vA, vB, vC
    Dim sForM As String
    Dim sCtrL As String
    Dim sCtrlStatus As String
    
    Dim lngLckdStatusBackColor As Long
    Dim lngLckdStatusForeColor As Long
    Dim lngLckdStatusBorderColor As Long
    Dim lngLckdStatusBorderStyle As Long
    Dim iLckdStatusSpecialEffect As Long


    sForM = fForM.Name
    If Not cTweakableCtrL Is Nothing Then sCtrL = cTweakableCtrL.Name

'MsgBox "teste - aplicando Enable/Disable para: [ " & sCtrl & " ]"
'Stop

    sCtrlStatus = IIf(bEnable = True, "Enbld", "Dsbld")

'Stop
    
    ' a habilita��o/desabilita��o � feita com base no indicador [ sCtrlType ] pois deve ser diferente a depender do tipo de controle
    '  . acCheckBox
    '  . acOptionGroup
    '
    '  . acTextBox
    '  . acListBox
    '  . acComboBox
    '
    '  . acCommandButton
'Stop
    
    'O par�metro [ bChangeEnable ] foi descartado logo ap�s a inclus�o do campo [ FormMode ]
    ' o [ FormMode ] substituiu a fun��o que tinha o [ bChangeEnable ] e ainda abriu a possibilidade de incluir
    ' outros [ FormModes ]
    
    

    'apenas esses tipos de controle podem ser alvo de Enable/Disable
    Select Case sCtrlType
        Case "btn", "chk", "opb", "txt", "lst", "cmb"
            Set clObjLckdStatusParam = dictCtrlEnblDsblParams(sForM)(sCtrL).dictParamByLckdStatus(sCtrlStatus)
        
            'recupera no Dict [ dictCtrlEnblDsblParams(sForM) ] os valores para os par�metros do controle
            lngLckdStatusBackColor = clObjLckdStatusParam.lngLckdStatusBackColor
            lngLckdStatusForeColor = clObjLckdStatusParam.lngLckdStatusForeColor
            lngLckdStatusBorderColor = clObjLckdStatusParam.lngLckdStatusBorderColor
            lngLckdStatusBorderStyle = clObjLckdStatusParam.lngLckdStatusBorderStyle
            iLckdStatusSpecialEffect = clObjLckdStatusParam.iLckdStatusSpecialEffect
        
            'A tentativa de ocultar um bot�o habilitado em com foco d� erro 2165
            ' desabilitar o bot�o antes dele ocult�-lo for�a a perda do foco e evita o erro
            If Not bVisible Then bEnable = bVisible
        
            'Come�a a altera��o do controle de Enabled pra Disabled e vice-versa
            With cTweakableCtrL
'Stop
                '------------------------------------------------
                'par�metros Back, Border e ForeColor
                ' apenas pra Text, List e Combobox
                '------------------------------------------------
                If sCtrlType <> "chk" And sCtrlType <> "opb" Then
                    'identifica o tipo de controle
                    'a cor de controle desabilitado para Bot�es e para Textos ou Listas � diferente
'                        sHexColor = IIf(sCtrlType = "1", sgbBtnGREyBackColor, IIf(sCtrlType = "2", sgbTxtBRownDsablBackColor, "#0"))
'                            Debug.Print sHexColor
'Stop
                    .BackColor = lngLckdStatusBackColor
                    .BorderColor = lngLckdStatusBorderColor
                    .ForeColor = lngLckdStatusForeColor
'Stop
                    '--------------------------------------------
                    'par�metro Gradient, apenas pra CommandButton
                    '--------------------------------------------
                    If sCtrlType = "btn" Then
                        .Gradient = 0
                        '.Gradient = 25
                    
                    End If
                
                    '--------------------------------------------
                    'par�metro
                    ' apenas pra Check e OptionButton
                    '--------------------------------------------
                End If
                    
                '------------------------------------------------
                'par�metros SpecialEffect e BorderStyle
                ' pra todos exceto CommandButton
                '------------------------------------------------
                If sCtrlType <> "btn" Then
                
                    .BorderStyle = lngLckdStatusBorderStyle
                    .SpecialEffect = iLckdStatusSpecialEffect
                
                End If
                    
                '----------------------------------------------------
                'par�metros Enabled/Locked
                ' se for um Check, OptionGrp, CommandButton ou Combo p�e Enabled como False, se for outro tipo de controle p�e Locked como True
                '----------------------------------------------------
'Stop
                If Not IsEmpty(bEnable) Then
                    If sCtrlType = "opt" Or sCtrlType = "btn" Then
                        .Enabled = bEnable
                        
                    
                    ElseIf sCtrlType = "cmb" Then
                        .Enabled = bEnable
                        
                        
'parei aqui: melhorar a explica��o
                        'H� [ Comboboxes ] como a [ cmbEdtSetor ] que ter�o que exibir
                        'Para Combos com multiplos valores pode-se usar este campo pra fazer com que ao habilitar
                        ' uma Combo ela fique bloqueada, evitando que usu�rio selecione itens, j� que neste caso a edi��o n�o ser� na combo mas sim num pop-up auxliar
                        .Locked = IIf(bLocked, bLocked, Not bEnable)
                        '.Locked = Not bEnable
                        
                    Else
                        .Locked = Not bEnable
                     
                    End If
                    
                End If
'Stop

'MsgBox "teste - OnError para: [ " & sCtrL & " ]"
'Stop
                
                On Error GoTo -1 ': On Error GoTo 0
                On Error Resume Next
                .Visible = bVisible
                If Err = 2165 Then
Stop
                    
                End If
                
            End With

    End Select
    
    cTweakableCtrL.ControlTipText = sCtrlNewTipText & Chr(160) & Chr(160)

End Sub
