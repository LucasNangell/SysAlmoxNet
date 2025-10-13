Attribute VB_Name = "M�dulo 01a - OpemForms"
Option Compare Database
Option Explicit


Public Sub FormLoad00a_FindSysPaths(sForM As String)
    
    '------------------------------------------------------------------------------------
    'Essa fun��o pra confirmar se os caminhos de arquivo usados pelo sistema est�o acess�veis
    ' poderia ser chamada apenas no Start do sistema
    ' Mas para que fosse poss�vel incluir eventuais erros no log de carga de formul�rio
    ' optou-se por cham�-la ao iniciar a abertura de cada formul�rio
    
    ' a chamada para a sub deve ser inserida no c�digo da sub [ FormLoad05_OpenForm ] do m�dulo [ M�dulo 01a - OpemForms ]
    ' Call FormLoad00a_FindSysPaths(sTrgtForm)
    '------------------------------------------------------------------------------------
    
    Dim vA
    Dim objVbComp As Object
    Dim iLcod As Integer
    Dim sVariavel As String
    Dim sLoadLogWarn As String
    Dim sPath As String, sPath2 As String, sPathFull As String
    Dim sFile As String
    Dim vCodMod As Variant
    Dim vPath As Variant
    Dim vKey As Variant
    
    
    
    Set objVbComp = Application.VBE.ActiveVBProject.VBComponents("M�dulo 00b - Vari�veis")
    
    'Atribui a [ vCodMod ] as linhas de c�digo do [ objVbComp ]
    vCodMod = Split(objVbComp.CodeModule.Lines(1, objVbComp.CodeModule.CountOfLines), vbCrLf)
    
    'Percorre cada linha de c�digo em [ vCodMod ]
    For iLcod = 0 To UBound(vCodMod)
        
        'Remove os espa�os a esquerda da linha ora analisada
        Do While Left(vCodMod(iLcod), 1) = " ": vCodMod(iLcod) = Right(vCodMod(iLcod), Len(vCodMod(iLcod)) - 1): Loop
        
        'Verifica se n�o se trata de uma linha comentada
        If Not Left(vCodMod(iLcod), 1) = "'" Then
            'Verifica se a linha atual possui "\\" ou ":\" o que indica um caminho
            If InStr(vCodMod(iLcod), "\\") > 0 Or InStr(vCodMod(iLcod), ":\") > 0 Then
                'Verifica a exist�ncia do sinal "=", o que caracteriza a atribui��o do caminho a alguma vari�vel
                If InStr(vCodMod(iLcod), " = ") > 0 Then
                    'Atribui o nome da vari�vel � [ sVariavel ]
                    sVariavel = Split(vCodMod(iLcod), " = ")(0)
                    sPath = Split(vCodMod(iLcod), " = ")(1)
                    sVariavel = Split(sVariavel, " As ")(0)
                    vA = Split(sVariavel, " ")
                    sVariavel = vA(UBound(vA))
                    'Remove aspas excessivas
                    sPath = Replace(sPath, """", "")
                    'Adiciona a [ sVariavel ] ao [ dictFndFldrVars ] com seu caminho [ sPath ]
                    If Not dictFndFldrVars.Exists(sVariavel) Then dictFndFldrVars.Add sVariavel, sPath
                    
                End If
            End If
            'Verifica se a linha do c�digo ora analisada cont�m alguma vari�vel de caminho do [ dictFndFldrVars ]
            For Each vKey In dictFndFldrVars
                If InStr(vCodMod(iLcod), vKey) > 0 And InStr(vCodMod(iLcod), " = ") > 0 Then
                    sVariavel = Replace(Replace(Split(vCodMod(iLcod), " = ")(0), "Public Const ", ""), " As String", "")
                    vPath = Split(vCodMod(iLcod), " = ")
                    sPath = Replace(vPath(1), """", "")
                    sPathFull = sPath
                    
                    'Verifica se o valor de [ sPath ] possui algum "&" o que sinaliza concatena��o de valores
                    If InStr(sPath, "&") > 0 Then
                        
                        'Caso haja o "&" separa o valor em 2 partes, sendo a primeira parte o caminho raiz, e a segunda parte o complemento
                        sPath = Replace(Split(vPath(1), " & ")(0), """", "")
                        sPath2 = Replace(Split(vPath(1), " & ")(1), """", "")
                        
                        'Busca o valor do caminho raiz para realizar a concatena��o
                        If dictFndFldrVars.Exists(sPath) Then
                            sPathFull = dictFndFldrVars(sPath) & sPath2
                        End If
                    End If
                    
                    'Adiciona a [ sVariavel ] ao dicion�rio com o caminho j� concatenado
                    If Not dictFndFldrVars.Exists(sVariavel) Then dictFndFldrVars.Add sVariavel, sPathFull
                    
                    'Verifica se o caminho existe e est� acess�vel ao sistema, caso n�o esteja carrega no log do sistema
                    If Dir(sPathFull) = "" Then
                        If InStr(sPathFull, ".") > 0 Then
                            sFile = Split(sPathFull, "\")(UBound(Split(sPathFull, "\")))
                            sPathFull = Replace(sPathFull, sFile, "")
                            
                            sLoadLogWarn = "O arquivo [ " & sFile & " ] indicado pela vari�vel [ " & sVariavel & " ] n�o est� acess�vel ou n�o existe no caminho:" & vbCrLf & sPathFull
                            Call FormStatusBar01_Bld(sForM, sVariavel, sLoadLogWarn)
                            dictFndFldrVars(sVariavel) = ""
                        Else
                            
                            sLoadLogWarn = "O caminho indicado pela vari�vel [ " & sVariavel & " ] n�o est� acess�vel ou n�o existe:" & vbCrLf & sPathFull
                            Call FormStatusBar01_Bld(sForM, sVariavel, sLoadLogWarn)
                            dictFndFldrVars(sVariavel) = ""
                        End If
                    End If
                End If
            Next vKey
        End If
    Next iLcod

End Sub


Sub FormLoad01_FormLoadingStart(clObjFormOpenParams As cls_09cParamsToOpenForms)
    Dim vA, vB
    Dim vFormCoords(1) As Variant
    Dim cTglBtnDocking As Control
'Stop
    
If gBbDepurandoLv02a Then MsgBox "----- FormLoad01_FormLoadingStart ---------------------------------------------" & vbCr & vbCr & "Confirma se [ " & clObjFormOpenParams.sTrgtForm & " ] foi aberto manualmente" & vbCr & "por um [ TriggerForm ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then Stop
    
    'Se houver um [ TriggForm ] que disparou a abertura do [ TrgtForm ]
    ' chama rotina pra recuperar dados do [ TrggForm ]
    If clObjFormOpenParams.sTrggForM <> "" Then
        
If gBbDepurandoLv02a Then MsgBox "----- FormLoad01_FormLoadingStart ---------------------------------------------" & vbCr & vbCr & "[ " & clObjFormOpenParams.sTrgtForm & " ] foi aberto pelo usu�rio" & vbCr & "por um [ TriggerForm ]" & vbCr & "chama [ FormLoad02a_SetFormPositionDict ]" & vbCr & " "
If gBbDepurandoLv02a Then Stop
        
        Call FormLoad02a_SetFormPositionDict(clObjFormOpenParams)
'Stop
    Else
'Stop
        
        
    End If
    
'Stop
    
If gBbDepurandoLv02a Then MsgBox "----- FormLoad01_FormLoadingStart ---------------------------------------------" & vbCr & vbCr & "Chama [ FormLoad05_OpenForm ] pra iniciar a abertura de" & vbCr & "[ " & clObjFormOpenParams.sTrgtForm & " ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then Stop
    
    'Inicia a abertura do [ TrgtForm ] usando os par�metros passados pela fun��o
    Call FormLoad05_OpenForm(clObjFormOpenParams)
    If clObjFormOpenParams.fTrgtForm Is Nothing Then Set clObjFormOpenParams.fTrgtForm = Forms(clObjFormOpenParams.sTrgtForm)
    If ControlExists("tglDockingTrggFrm", clObjFormOpenParams.fTrgtForm) Then
        Set clObjFormOpenParams.cTglBtnDocking = Forms(clObjFormOpenParams.sTrgtForm).Controls("tglDockingTrggFrm")
        Call FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams)
    End If
'Stop
    Call Scr_FormAlwaysOnTop(clObjFormOpenParams.fTrgtForm, False)
End Sub


Public Sub FormLoad02a_SetFormPositionDict(clObjFormOpenParams As cls_09cParamsToOpenForms)
               
    Dim vA, vB, vC
    Dim fTrggForm As Form
    Dim lngTrgtLeft As Long
    Dim lngTrgtTop As Long
    Dim lngTrgtHorizOffset As Long
    Dim lngTrgtVertOffset As Long

    Set fTrggForm = Forms(clObjFormOpenParams.sTrggForM)

    
'MsgBox "FormLoad02a_SetFormPositionDict ------------------------------------------------" & vbCr & "load dictFormsParams"
'Stop
    'Monta o dicion�rio com par�metros do [ sTrggForm ]
    ' .coordenadas de posi��o
    ' . ...
    If dictFormsParams.Exists(clObjFormOpenParams.sTrggForM) Then
        Set clObjFormsParams = dictFormsParams(clObjFormOpenParams.sTrggForM)

        lngTrgtHorizOffset = clObjFormsParams.lngTrgtHorizOffset
        lngTrgtVertOffset = clObjFormsParams.lngTrgtVertOffset

    Else
        Set clObjFormsParams = New cls_09aFormsParams
        dictFormsParams.Add clObjFormOpenParams.sTrggForM, clObjFormsParams
        
        lngTrgtHorizOffset = fTrggForm.Width + GbLngTrgtHorizOffset
        lngTrgtVertOffset = GbLngTrgtVertOffset

    End If
        
    clObjFormsParams.sTrggFormName = clObjFormOpenParams.sTrggForM
    clObjFormsParams.sTrgtFormName = clObjFormOpenParams.sTrgtForm
    clObjFormsParams.lngTrggLeft = fTrggForm.WindowLeft
    clObjFormsParams.lngTrggTop = fTrggForm.WindowTop
    clObjFormsParams.lngTrgtHorizOffset = lngTrgtHorizOffset
    clObjFormsParams.lngTrgtVertOffset = GbLngTrgtVertOffset
    clObjFormsParams.lngTrgtLeft = fTrggForm.WindowLeft + lngTrgtHorizOffset
    clObjFormsParams.lngTrgtTop = fTrggForm.WindowTop + GbLngTrgtVertOffset
    clObjFormsParams.bTrgtFormIsDocked = clObjFormOpenParams.bFrmIsDocked

End Sub


Public Sub FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams As cls_09cParamsToOpenForms)
    Dim vA, vB, vC
    Dim iInT As Integer
    Dim cTglDockingTrggFrm As Control
    Dim cTglDockingTrgtFrm As Control
    Dim bTrggFormIsOpen As Boolean
    Dim bTrgtFormIsOpen As Boolean
    Dim sTipText As String

    'Verificar se o formul�rio que acionou o bot�o � um [ TrggForm ] ou [ TrgtForm ]
    '   a verifica��o dever� ser feita atrav�s do nome do bot�o
    With clObjFormOpenParams
    'Verifica se o bot�o ativo � um [ tglDocking ]
    '  caso seja [tglDockingTrggFrm] indica que foi pressionado em um [TrgtForm]
    '  caso seja [tglDockingTrgtFrm] indica que foi pressionado em um [TrggForm]
        If InStr(1, .cTglBtnDocking.Name, "tglDocking", vbTextCompare) > 0 Then
            If .cTglBtnDocking.Name = "tglDockingTrggFrm" Then .bIsTrggForm = False Else .bIsTrggForm = True
        Else
            'caso n�o seja um [ tglDocking ] a vari�vel [ bIsTrggForm ] recebe o valor a depender
            '  do [ dictFormsParams ], se existir um dicion�rio com o [ sForm ] indica um [ TrggForm ]
            .bIsTrggForm = dictFormsParams.Exists(.sTrgtForm)
        End If
    
        'Se o bot�o foi pressionado em um [ TrggForm ]
        If .bIsTrggForm Then
        '   se existir um dicion�rio respectivo pega o [ TrggForm ] e o [ TrgtForm ] relacionados
            If dictFormsParams.Exists(.sTrgtForm) = True Then
                'Verifica se o dicion�rio possui a classe com os par�metros
                If IsObject(dictFormsParams(.sTrgtForm)) Then
                    Set clObjFormsParams = dictFormsParams(.sTrgtForm)
                    .sTrggForM = clObjFormsParams.sTrggFormName
                    .sTrgtForm = clObjFormsParams.sTrgtFormName
                    If .sTrggForM = "" Then .sTrggForM = .sTrgtForm
                Else
                    'Caso o dicion�rio esteja vazio, apenas o bot�o do TrggForm ser� alterado
                    .sTrggForM = .sTrgtForm
                    .sTrgtForm = ""
                
                End If
            'Se o dicion�rio n�o existe, indica que apenas o [TrggForm] est� aberto
            '   sendo assim, apenas o formul�rio aberto sofrer� altera��es
            Else
                .sTrggForM = .sTrgtForm
                .sTrgtForm = ""
            End If
            
        'Caso o bot�o tenha sido acionado em um [TrgtForm]
        Else
            'Atribui a vari�vel [ sTrgtForm ] o valor de [ sForm ]
            .sTrgtForm = .sTrgtForm
            'Exclui valores possivelmente armazenados em [ sTrggForm ]
            .sTrggForM = ""
            
            'Faz uma varredura nos itens do dicion�rio para encontrar qual � o [TrggForm]
            For iInT = 0 To dictFormsParams.Count - 1
                If IsObject(dictFormsParams(dictFormsParams.Keys(iInT))) Then
                    Set clObjFormsParams = dictFormsParams(dictFormsParams.Keys(iInT))
                    'Se o par�metro [sTrgtFormName] for igual ao nome do formul�rio que o bot�o foi pressionado
                    '   a vari�vel [ sTrggForm ] recebe o valor armazenado na classe
                    If clObjFormsParams.sTrgtFormName = .sTrgtForm Then .sTrggForM = clObjFormsParams.sTrggFormName
                End If
            Next iInT
            
            
        End If
        
        'Verifica se existe um [ TrgtForm ]
        If .sTrgtForm <> "" Then
        '       Verifica se o [ TrgtForm ] est� aberto
            If CurrentProject.AllForms(.sTrgtForm).IsLoaded Then
                bTrgtFormIsOpen = True
                'Verifica a exist�ncia do bot�o no [ TrgtForm ]
                'Caso exista o bot�o seta a vari�vel [ cTglDockingTrggFrm ]
                If ControlExists("tglDockingTrggFrm", Forms(.sTrgtForm)) Then
                    Set cTglDockingTrggFrm = Forms(.sTrgtForm).Controls("tglDockingTrggFrm")
                Else
                    Set cTglDockingTrgtFrm = Nothing
                End If
            End If
            
        End If
        'Verifica se existe um [ TrggForm ]
        If .sTrggForM <> "" Then
        '       Verifica se o [ TrggForm ] est� aberto
            If CurrentProject.AllForms(.sTrggForM).IsLoaded Then
                bTrggFormIsOpen = True
                'Verifica a exist�ncia do bot�o no [ TrggForm ]
                'Caso exista o bot�o seta a vari�vel [ cTglDockingTrgtFrm ]
                If ControlExists("tglDockingTrgtFrm", Forms(.sTrggForM)) Then
                    Set cTglDockingTrgtFrm = Forms(.sTrggForM).Controls("tglDockingTrgtFrm")
                Else
                    Set cTglDockingTrgtFrm = Nothing
                End If
            End If
            
        End If
        
        'Se [ cTglDockingTrggFrm ] contiver algum controle modifica as propriedades do mesmo
        If Not cTglDockingTrggFrm Is Nothing Then
            sTipText = IIf(.bFrmIsDocked = True, " desancorar do form disparador", " ancorar com form disparador")
            'Verifica se os caminhos definidos pelas globais existem e est�o acess�veis
            If dictFndFldrVars("gBsSysLocalIcoTrggUnlock") <> "" And dictFndFldrVars("gBsSysLocalIcoTrggLock") <> "" Then
                cTglDockingTrggFrm.Picture = IIf(.bFrmIsDocked = True, gBsSysLocalIcoTrggUnlock, gBsSysLocalIcoTrggLock)
            End If
            cTglDockingTrggFrm.ControlTipText = sTipText
            cTglDockingTrggFrm.Value = .bFrmIsDocked
        End If
        
    'parei aqui1: erro ao acessar arquivo .png
        
        
        'Se [ cTglDockingTrgtFrm ] contiver algum controle modifica as propriedades do mesmo
        If Not cTglDockingTrgtFrm Is Nothing Then
            sTipText = IIf(.bFrmIsDocked = True, " desancorar do form alvo", " ancorar com form alvo")
            'Verifica se os caminhos definidos pelas globais existem e est�o acess�veis
            If dictFndFldrVars("gBsSysLocalIcoTrgtUnlock") <> "" And dictFndFldrVars("gBsSysLocalIcoTrgtLock") <> "" Then
                cTglDockingTrgtFrm.Picture = IIf(.bFrmIsDocked = True, gBsSysLocalIcoTrgtUnlock, gBsSysLocalIcoTrgtLock)
            End If
            cTglDockingTrgtFrm.ControlTipText = sTipText
            cTglDockingTrgtFrm.Value = .bFrmIsDocked
        End If
        
        'Atualiza as posi��es dos formul�rios na tela caso [ bFormIsDocked ] seja True
        
        'Verifica se existe um dicion�rio para atualizar
        If dictFormsParams.Exists(.sTrggForM) = True Then
            If IsObject(dictFormsParams(.sTrggForM)) Then
                Set clObjFormsParams = dictFormsParams(.sTrggForM)
                'Se o [ TrggForm ] estiver aberto atualiza as posi��es no dicion�rio
                If .sTrggForM <> "" Then
                    If CurrentProject.AllForms(.sTrggForM).IsLoaded Then
                        clObjFormsParams.lngTrggLeft = Forms(.sTrggForM).WindowLeft
                        clObjFormsParams.lngTrggTop = Forms(.sTrggForM).WindowTop
                    End If
                End If
                'Se o [ TrgtForm ] estiver aberto atualiza as posi��es no dicion�rio
                If .sTrgtForm <> "" Then
                    If CurrentProject.AllForms(.sTrgtForm).IsLoaded Then
                        clObjFormsParams.lngTrgtLeft = Forms(.sTrgtForm).WindowLeft
                        clObjFormsParams.lngTrgtTop = Forms(.sTrgtForm).WindowTop
                    End If
                End If
                'Atualiza o restante dos valores do dicion�rio caso [ bFormIsDocked ] seja verdadeiro
                If .bFrmIsDocked Then
                    clObjFormsParams.lngTrgtHorizOffset = clObjFormsParams.lngTrgtLeft - clObjFormsParams.lngTrggLeft
                    clObjFormsParams.lngTrgtVertOffset = clObjFormsParams.lngTrgtTop - clObjFormsParams.lngTrggTop
                End If
                
                clObjFormsParams.bTrgtFormIsDocked = .bFrmIsDocked
                
            End If
        
        End If
            
    End With
    
End Sub

Sub FormLoad05_OpenForm(clObjFormOpenParams As cls_09cParamsToOpenForms)

    Dim vA, vB, vC, vD
    'Dim sOpenArgs As String
    'Dim sSystemStartForm As String
    
    'Ap�s a carga dos Dicion�rios prossegue com abertura do formul�rio
    ' e carrega tempvars com par�metros de abertura pra n�o passar via OpenArgs
    '-----------------------------------
    
If gBbDepurandoLv02a Then MsgBox "----- FormLoad05_OpenForm  -----------------------------------------------------" & vbCr & vbCr & "Abre o form [ " & clObjFormOpenParams.sTrgtForm & " ]" & vbCr & "e inicia seu evento Load" & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then Stop

    Call FormLoad00a_FindSysPaths(clObjFormOpenParams.sTrgtForm)
'Stop
    If CurrentProject.AllForms(clObjFormOpenParams.sTrgtForm).IsLoaded Then DoCmd.Close acForm, clObjFormOpenParams.sTrgtForm
    DoCmd.OpenForm clObjFormOpenParams.sTrgtForm

    '-------------------------------------------------------------
    '-----------------------------------
    
        End Sub

Function FormLoad06b_BackFromFormLoad(sForM As String) As Boolean

    'No caso de haver no dict [ dictFormsParams(sForM) ] alertas levantados durante a carga do sistema
    ' monta o Texto a ser exibido na [ Status Bar ] do [ Form ]
    ' . a atualiza��o da [ Status Bar ] � feita na carga do formul�rio
    ' . s�o inclu�dos apenas alertas que n�o s�o tratados com o encerramento do sistema em outras partes da carga
    '   Ex.: inexist�ncia de TargtCtrls a serem filtrados, aus�ncia de par�metros, etc
    '------------------------------------------------------------------------
    'Confirma se o dict [ dictFormsParams ] existe
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "BackFromLoad"
'Stop
    If IsObject(dictFormsParams) Then
        
        'Confirma se no dict [ dictFormsParams ] foi inclu�do um item associado ao [ sForM ] com alertas a serem exibidos
        If dictFormsParams.Exists(sForM) Then
        
            Set clObjFormsParams = dictFormsParams(sForM)
            'Se a propriedade [ bForm_ShowWarns ] for True significa que
            ' h� alertas de falha na carga do sistema a serem exibidas no Form
            If clObjFormsParams.bForm_ShowWarns Then FormLoad06b_BackFromFormLoad = True Else FormLoad06b_BackFromFormLoad = False
        
        End If
    
    End If
    
End Function

Sub FormLoad06a_BackFromFormLoad(clObjFormOpenParams As cls_09cParamsToOpenForms)
    
    Dim vA, vB, vC, vD
    Dim cTargtCtrl As Control, cRecCntCtrL As Control
    Dim sTargtCtrl As String, sRecCntCtrl As String
    Dim sRecCntCtrlName As String
    Dim vKeyFilGrp As Variant, vKeyTrggCtrl As Variant, vKeyTrgtCtrl As Variant
    Dim vKey As Variant
    Dim sLoadLogWarn As String
    Dim cCtrL As Control
    Dim fForM As Form
    Dim sSQL As String
    Dim lngFilteredRecs As Long
    Dim sStR As String
    Dim bBoL As Boolean
    Dim sFilGrp As String
    Dim sFormMode As String
    Dim vOpenArgs As Variant
    Dim vFormCoords(1) As Variant
    Dim sQryIDfield As String, sQryIDfieldTrgt As String
    Dim cTrgtCtrlOrg As Control
    
'MsgBox "----- FormLoad06a_BackFromFormLoad ------------------------------------------" & vbCr & vbCr & "O evento Load do [ Form ] chama [ FormLoad06a_BackFromFormLoad ]" & vbCr & "pra configurar o [ Form ]." & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then Stop
'Stop
    
    
    If clObjFormOpenParams.bSetPosition Then
        
If gBbDepurandoLv02a Then MsgBox "----- FormLoad06a_BackFromFormLoad ------------------------------------------" & vbCr & vbCr & "chama [ FormLoad08_SetFormPosition ]" & vbCr & "pra definir a posi��o de abertura do [ Form ]." & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then Stop

        Call FormLoad08_SetFormPosition(clObjFormOpenParams)
'Stop
    End If

    'Se os dados do usu�rio tiverem sido recuperados na carga do sitema exibe no cabe�alho do formul�rio
    If ControlExists("lblUser", clObjFormOpenParams.fTrgtForm) Then
        If Not clObjUserParams Is Nothing Then clObjFormOpenParams.fTrgtForm.Controls("lblUser").Caption = clObjUserParams.sUserLogin & " - " & clObjUserParams.sUserName
    End If
'Stop
        'posi��o antiga da rotina de
        'Confirma��o se h� alertas de carga do sistema pra serem exibidos no [ Form ]

'Stop
    '------------------------------------------------------------------------
    '---------------------------------------
    'Verifica se o [ TrgtForm ] que est� sendo aberto possui valores padr�o pr�-definidos nos [ TrggCtrls ]
    ' pra isso passa por cada um dos Grupos de Filtragem do formul�rio e atualiza
    ' .os TargtCtrls
    ' .os respectivos RecCnts de cada TargCtrl

    'Se n�o houver Grupos de Filtragem no Form ou se o Form tiver sido aberto
    ' sem a sequ�ncia de carga de dicion�rios, desconsidera a atualiza��o
    bBoL = IsObject(dictFormFilterGrpTrgts(clObjFormOpenParams.sTrgtForm)) 'verifica se h� Grupos de Filtragem inclu�dos no Dicion�rio
Stop
    
    If bBoL Then
        For Each vKeyFilGrp In dictFormFilterGrpTrgts(clObjFormOpenParams.sTrgtForm)
            sFilGrp = vKeyFilGrp
            
            For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(clObjFormOpenParams.sTrgtForm)(vKeyFilGrp)
            
                '-----------------------------------------------------------
                'Atualiza o TargtCtrl e o RecCnt de cada Grupo de Filtragem que tenha [ TriggCtrls ] adicionados ao [ dictFormFilterGrpTrgts(sForM)(vKeyFilGrp) ]
                '-----------------------------------------------------------
                Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(clObjFormOpenParams.sTrgtForm)(vKeyFilGrp)(vKeyTrgtCtrl)
                sTargtCtrl = clObjTargtCtrlParam.sTargtCtrlName
                sRecCntCtrl = clObjTargtCtrlParam.sRecCntCtrlName
                
                On Error GoTo -1
                Call pb_TargtCtrlUpdate06_BuildWHERE(clObjFormOpenParams.fTrgtForm, sFilGrp)
                
                'Caso [ bShowRecID ] seja verdadeira indica que o formul�rio que est� sendo aberto deve exibir um registro espec�fico
                If clObjFormOpenParams.bShowRecID Then
    
                    'Recupera o [ sQryIDfield ] do [ TrgtCtrl ] do [ TrgtForm ] para compara��o com o [ clObjFormOpenParams.sQryFieldID ]
                    ' afim de definir qual [ TrgtCtrl ] dever� ser afetado com o [ lngRecID ] informado
                    sQryIDfieldTrgt = GetsQryFieldID(clObjTargtCtrlParam)
                    'Caso [ sQryIDField ] seja igual a [ sQryIDFieldTrgt ] indica que o controle destino foi encontrado
                    ' ent�o define [ cTargtCtrl ] com o controle ora avaliado
                    If clObjFormOpenParams.sQryFieldID = sQryIDfieldTrgt Then
                        Set cTargtCtrl = Forms(clObjFormOpenParams.sTrgtForm).Controls(sTargtCtrl)
    
                            cTargtCtrl.Selected(clObjFormOpenParams.lngRecID) = True
                            cTargtCtrl.ListIndex = clObjFormOpenParams.lngRecID
                    
                            Call PbSubFillFieldsByList(cTargtCtrl)
                    End If
                    
                End If
            
            Next vKeyTrgtCtrl
            
        Next vKeyFilGrp

    End If

    '---------------------------------------
    '------------------------------------------------------------------------
    'S� chama a rotina se tiver sido feita a carga completa do sistema,
    ' incluindo a carga dos dicion�rios

    If gBbDictsLoaded Then
        Call pbSub00_CtrlsEnblDsble_GetParams(clObjFormOpenParams.fTrgtForm, clObjFormOpenParams.sFormMode)
    
        'Confirma se a inicializa��o de [ Eventos de Classe ] est� ativada na vari�vel global
        If gBbInitCtrlEvents Then
            
            'Chama rotina pra montar o dict [ dictCtrlsEvents(sForM) ]
            If gBbDebugOn Then Debug.Print "Ctrl Events dict init"
            Call FormLoad07_GenCtrlsEventDictInit(clObjFormOpenParams.sTrgtForm)
        Else
            sLoadLogWarn = "A inicializa��o de [ Eventos de Classe ] est� desativada." & vbCrLf & "CtrlsBehvr e pesquisas n�o ir�o funcionar."
            Call FormStatusBar01_Bld(clObjFormOpenParams.sTrgtForm, "InitEvents", sLoadLogWarn)
        
        End If
    
    End If
    
vA = "----- FormLoad06a_BackFromFormLoad ------------------------------------------" & vbCr & vbCr & "chama [ FormStatusBar02_SetWarn(fForM) ]" & vbCr
vB = "pra verificar se h� alertas de carga do sistema pra serem exibidas na Barra de Status." & vbCr & " " & vbCr & " "
If gBbDepurandoLv02a Then MsgBox vA & vB
If gBbDepurandoLv02a Then Stop
    
    If ControlExists("lblStatusBar", clObjFormOpenParams.fTrgtForm) Then
        
        'Confirma se h� alertas de carga do sistema pra serem exibidos no [ Form ]
        bBoL = FormLoad06b_BackFromFormLoad(clObjFormOpenParams.sTrgtForm)
        Call FormStatusBar02_SetWarn(clObjFormOpenParams.fTrgtForm, bBoL)
        'If FormLoad06b_BackFromFormLoad(clObjFormOpenParams.sTrgtForm) Then Call FormStatusBar02_SetWarn(clObjFormOpenParams.fTrgtForm)
        bBoL = False
    
    End If
    
    'Chama a rotina pra ocultar o Access e tornar vis�vel o Formul�rio principal
    ' essa a��o for�a a abertura do form antes que o evento Load do pr�prio form seja conclu�do
    'Call Scr_FormAlwaysOnTop(clObjFormOpenParams.fTrgtForm, False)

End Sub


Public Sub FormLoad07_GenCtrlsEventDictInit(sForM As String)
    
    Dim vA, vB
    Dim cCtrL As Control
    Dim vKeyEventCtrls As Variant, vKeyFilGrps As Variant, vKeyTrgtCtrl As Variant
    Dim sTrgtCtrl As String
    
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Init Gen Ctrls events [ " & sForM & "]"
'Stop
'parei aqui:
    
    'Toda vez que um [ form ] for aberto passa por essa rotina pra
    ' inicializar os controles gen�ricos, inclu�dos no dict [ dictCtrlsEvents(sForm) ]
    If Not IsObject(dictCtrlsEvents(sForM)) Then Exit Sub
    For Each vKeyEventCtrls In dictCtrlsEvents(sForM)
        
        Set cCtrL = Forms(sForM).Controls(vKeyEventCtrls)
        Set dictCtrlsEvents(sForM)(vKeyEventCtrls).InitCtrl = cCtrL
    
    Next vKeyEventCtrls
    
'Stop
    
    'Inicializa os controles tipo [ DataField ] inclu�dos no dict [ dictFormFilterGrpTrgts(sForm) ]
    ' pra automatizar a exibi��o de dados do [ Listbox ]
    
    If Not IsObject(dictFormFilterGrpTrgts(sForM)) Then Exit Sub
    
    For Each vKeyFilGrps In dictFormFilterGrpTrgts(sForM)
        For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(sForM)(vKeyFilGrps)
        
            Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(vKeyFilGrps)(vKeyTrgtCtrl)
            sTrgtCtrl = clObjTargtCtrlParam.sTargtCtrlName
            
    'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Init TrgtCtrls events [ " & sTrgtCtrl & "]"
    'Stop
            
            Set cCtrL = Forms(sForM).Controls(sTrgtCtrl)
            If cCtrL.ControlType = acListBox Then Set dictFormFilterGrpTrgts(sForM)(vKeyFilGrps)(vKeyTrgtCtrl).InitCtrl = cCtrL
        Next vKeyTrgtCtrl
    Next vKeyFilGrps

End Sub


Sub FormLoad08_SetFormPosition(clObjFormOpenParams As cls_09cParamsToOpenForms)
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "SetFormPosition"
'Stop
    
    'Para formul�rios que est�o sendo abertos define a posi��o inicial
    ' para formul�rios que j� est�o abertos faz o reposicionamento a partir dos par�metros de ancoragem
    With clObjFormOpenParams

        If .sTrggForM <> "" Then
            If .bSetPosition And CurrentProject.AllForms(.sTrggForM).IsLoaded Then
                .lngFormLeft = Forms(.sTrggForM).WindowLeft + .lngFormLeft
                .lngFormTop = Forms(.sTrggForM).WindowTop + .lngFormTop
                
                If .bCentralizeForm Then
                    .lngFormLeft = Forms(.sTrggForM).WindowLeft + (Forms(.sTrggForM).InsideWidth / 2) - (Forms(.sTrgtForm).InsideWidth / 2)
                    .lngFormTop = Forms(.sTrggForM).WindowTop + (Forms(.sTrggForM).InsideHeight / 2) - (Forms(.sTrgtForm).InsideHeight / 2)
                End If
            Else
                If dictFormsParams.Exists(.sTrggForM) Then
                    Set clObjFormsParams = dictFormsParams(.sTrggForM)
                    
                    If .bIsTrggForm Then
                        .lngFormLeft = clObjFormsParams.lngTrggLeft
                        .lngFormTop = clObjFormsParams.lngTrggTop
                    Else
                        .lngFormLeft = clObjFormsParams.lngTrgtLeft
                        .lngFormTop = clObjFormsParams.lngTrgtTop
                        
                    End If
                End If
            End If

        End If
        
        .fTrgtForm.Move .lngFormLeft, .lngFormTop
    End With
'Stop
End Sub

Sub SysLoad00_SysStartLoad()

    Dim vA, vB, vC
    Dim sTrgtForm As String
    Dim sFormMode As String
    Dim lngTrggRecID As Long
    Dim bSetTrgtPos As Boolean
    Dim bTrgtFormIsDocked As Boolean
    Dim sTrggForM As String
    Dim cTglDockingFrm As Control
'Stop
    'Recupera o nome do [ formul�rio principal ]
    sTrgtForm = DLookup("[StartForm]", "tbl_01(01)aSysStart", "[SysParamsID]= 1")
'Stop
    'Atribui��o dos valores a [ clObjFormOpenParams ] usados nas rotinas de abertura de formul�rios
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams

        'If Not IsNull(cCtrLDockingTrgtFrm.Value) Then Set .cTglBtnDocking = cCtrLDockingTrgtFrm
        'If Not .cTglBtnDocking Is Nothing Then .bFrmIsDocked = .cTglBtnDocking.Value
        
        .bIsTrggForm = False
        .bSetPosition = True
        '.fTrgtForm   este valor deve ser definido ap�s a abertura do formul�rio em FormLoad05_OpenForm
        .lngFormLeft = GbLngTrgtLeft
        .lngFormTop = GbLngTrgtTop
        '.lngRecID = 30
        .sFormMode = "StartView"
        .sTrgtForm = sTrgtForm
        .sTrggForM = ""
        .sTrgtForm = DLookup("[StartForm]", "tbl_01(01)aSysStart", "[SysParamsID]= 1")
    
    End With
If gBbDepurandoLv01a Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Sys Start"
If gBbDepurandoLv01a Then Stop

    'Faz a carga completa dos dicion�rios do sistema
    Call SysLoad01_SysDictsLoad(sTrgtForm)
'Stop

If gBbDepurandoLv01a Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Back from SysDict Load"
If gBbDepurandoLv01a Then Stop
    
'Stop
'    'Par�metros pra abertura do [ TrgtForm ]
'    sTrgtForm = sTrgtForm
'    sFormMode = "ProdView"
'    bTrgtFormIsDocked = Me.cCtrLDockingTrgtFrm.Value
'    bSetTrgtPos = True
'
    

If gBbDepurandoLv01a Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Inicia as rotinas pra abertura" & vbCr & "do form [ " & sTrgtForm & " ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
    
    'Inicia as rotinas pra abertura de formul�rio
'    On Error GoTo -1


    
    Call FormLoad01_FormLoadingStart(clObjFormOpenParams)
    'Call FormLoad01_FormLoadingStart(sTrgtForm, sFormMode, bSetTrgtPos, bTrgtFormIsDocked, sTrggForm, lngTrggRecID)

End Sub


Sub SysLoad01_SysDictsLoad(sSystemStartForm)
    Dim vA, vB, vC
    Dim oFrmObjct As Object     'Formul�rio a ser carregado
    Dim sForM As String         'Nome do formul�rio a ser carregado
    Dim fForM As Form
    Dim vKey As Variant
    Dim vLoginStR As Variant
    Dim sLoadLogWarn As String
    
' Stop
    '-------------------------------------------------------------
    '-----------------------------------
    'Checagens iniciais
    ' - verifica se o sistema est� bloqueado para uso devido a manuten��o
    ' - interrup��o for�ada de clientes (de acordo com timer)
    ' - verifica atualiza��o de vers�o (atualizar automaticamente)
    ' - verificar se a pasta local do sistema com �cones de bot�es entre outros, foi encontrada
    '-----------------------------------
    '-------------------------------------------------------------
    
    
    'Chama rotina pra garantir que todos os dicion�rios com par�metros de controles do sistema
    ' estejam vazios antes de inicar a rotina de montagem de dicion�rios
    '-------------------------------------------------------------
    '-----------------------------------
'Stop
    Call CleanDicts
'Stop
    'Carrega o dicion�rio [ dictCtrlTypeStR ], contendo o nome por extenso dos tipos de controle
    ' que ser� concatenado com "BuildSQL_"  pra chamar as fun��es de Filtragem
    dictCtrlTypeStR.Add 104, "CommButton"
    dictCtrlTypeStR.Add 105, "Option"
    dictCtrlTypeStR.Add 106, "CheckBox"
    dictCtrlTypeStR.Add 107, "OptionGroup"
    dictCtrlTypeStR.Add 109, "TextBox"
    dictCtrlTypeStR.Add 110, "ListBox"
    dictCtrlTypeStR.Add 111, "ComboBox"
    
    'Carrega o dicion�rio [ dictCtrlTypeShort ], contendo a abreviatura dos tipos de controle
    ' pra facilitar a identifica��o do controle em trechos do sistema
    dictCtrlTypeShort.Add 104, "btn"
    dictCtrlTypeShort.Add 105, "opb"
    dictCtrlTypeShort.Add 106, "chk"
    dictCtrlTypeShort.Add 107, "grp"
    dictCtrlTypeShort.Add 109, "txt"
    dictCtrlTypeShort.Add 110, "lst"
    dictCtrlTypeShort.Add 111, "cmb"
    
    
'    dictCtrlType.Add "104", "CommButton"
'    dictCtrlType.Add "106", "CheckBox"
'    dictCtrlType.Add "107", "OptionGroup"
'    dictCtrlType.Add "109", "TextBox"
'    dictCtrlType.Add "110", "ListBox"
'    dictCtrlType.Add "111", "ComboBox"
    '-----------------------------------
    '-------------------------------------------------------------
    
'Stop
    
    
    gBbDictsLoaded = True
    vLoginStR = Environ("username")
    'vLoginStR = "6320"  'ao final do desenvolvimento remover a linha


'GoTo SkipTo2

    '-----------------------------------------------------------------------
    'Carrega o dicion�rio com as permiss�es de acesso do usu�rio logado
'Stop
    Call pbSub00_UserPermissionsDictBuild(vLoginStR)
    '-----------------------------------------------------------------------
    
'Stop

    '-----------------------------------------------------------------------
    'Confirma que a pasta de sistema foi localizada
    '-----------------------------------------------------------------------
'Stop
SkipTo1:
    '-----------------------------------
    '-------------------------------------------------------------
    'Percorre cada um dos Forms, abre em modo oculto e carrega os dicion�rios do sistema
     For Each oFrmObjct In CurrentProject.AllForms
        sForM = oFrmObjct.Name
'Stop
        
If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Abre o  form [ " & sForM & " ] e inicia a" & vbCr & "carga dos dicion�rios" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
        
        'S� faz a carga dos controles se o formul�rio n�o for o SysStart
        If sForM <> gBsSystemDefaultForm Then
'Stop
            'S� faz a carga dos dicion�rios se o formul�rio ora avaliado for o que est� indicado como principal na tabela SysStart
            ' parei aqui1: apagar quando sistema estiver finalizado
            If sForM <> sSystemStartForm And sForM <> "frm_02(1)aProdCadastro" And sForM <> "frm_00(1)cSysLoadLog" Then GoTo SKIP_ALL
        
'Stop
'GoTo SkipTo4
            
            'acDesign evita que o c�digo de abertura do formul�rio seja executado
            DoCmd.OpenForm sForM, acDesign, , , , acHidden
            Set fForM = Forms(sForM)


'If gBbDepurandoLv01a Then MsgBox "teste --------------------------------------------------------------------------"
If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Abre o  form [ " & sForM & " ] e inicia a" & vbCr & "carga dos dicion�rios" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop

''GoTo SkipTo2

        'inicia a montagem de dicion�rios do Formul�rio ora avaliado
        '-------------------------------------------------------------
        '-----------------------------------
            
            If gBbDebugOn Then Debug.Print "------" & vbCr & sForM & "--"
            If gBbDebugOn Then Debug.Print "Form Controls get parameters"

'GoTo SkipTo4
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "start TargtCtrlsDict"
'Stop
'Stop

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Chama [ pbSub20_TargtCtrlsDictStartUp ] pra montagem do" & vbCr & "dict [ dictFormFilterGrpTrgts(sForm) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop

            'Chama a rotina pra iniciar a montagem do dicion�rio de [ TargtCtrls ]
            ' e retorna se foram encontrados [ TargtCtrls ]
            '-------------------------------------------------------------
            '-----------------------------------
            gBbTrgtCtrlsFound = pbSub20_TargtCtrlsDictStartUp(fForM)
            '-----------------------------------
            '-------------------------------------------------------------


SkipTo2:

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub20_TargtCtrlsDictStartUp ] ap�s" & vbCr & "avaliar todos os controles" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
            
If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Avalia [ aus�ncia de TargtCtrls ] pra incluir na [ StatusBar ] do [ Form ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
            'Se n�o houver [ TargtCtrls ] no Form usa a barra de status do formul�rio pra alertar o usu�rio
            If gBbDebugOn Then Debug.Print "Trigger Controls"

            If Not gBbTrgtCtrlsFound Then
                

vA = "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & "chama [ FormStatusBar01_Bld ] e inclui o erro de" & vbCr & "[ aus�ncia de TargtCtrls ] na [ StatusBar ] de [" & Chr(160) & sForM & Chr(160) & "]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then MsgBox vA
If gBbDepurandoLv01a Then Stop
                sLoadLogWarn = "N�o foram encontrados TargetCtrls no formul�rio. N�o ser� poss�vel fazer pesquisas."
                Call FormStatusBar01_Bld(sForM, "NoTrgtCtrls", sLoadLogWarn)
            Else


If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "N�o houve erro de aus�ncia TargtCtrl pra inclus�o no" & vbCr & "dict de log de erros" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
            
            End If
            

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Chama [ pbSub30_TriggCtrlDictStartUp ] pra montagem dos" & vbCr & "dicts [ dictTrgg00GrpsInForm ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
            'Chama rotina pra iniciar a montagem do dicion�rio [ CtrlsBehvrParams ]
            ' e, se houverem sido localizados [ TargtCtrls ] no Form, tamb�m o dicion�rio de [ TriggCtrls ]
            '-------------------------------------------------------------
            '-----------------------------------
            Call pbSub30_TriggCtrlDictStartUp(fForM)  'Call pbSub10_EventsDictBuild(sForM, sCtrL) chamado internamente

            '-----------------------------------
            '-------------------------------------------------------------

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub30_TriggCtrlDictStartUp ] e" & vbCr & "[ pbSub41_CtrlsBehvrDictBuild ] ap�s avaliar todos os controles" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop


'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "back from TriggCtrlDictStartUp"
'Stop

'GoTo SkipTo4
            'Debug.Print "ResetAreas Controls"
            
GoTo SkipTo4

'If gBbDepurandoLv01b Then MsgBox "teste - back from RstAreaDict ------------------------------------------------" & vbCr & ""
'If gBbDepurandoLv01b Then Stop

'GoTo SkipTo4
SkipTo4:


'MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Form [ " & sForm & " ] " & "Chama [ pbSub60_CtrlsEnblDsblDictStartUp ] pra montagem do" & vbCr & "dict [ dictCtrlEnblDsblParams(sForM) ]" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
'Stop
            'Chama rotina pra iniciar a montagem do dicion�rio de [ ctrls Enble/Dsble ]
            '-------------------------------------------------------------
            '-----------------------------------
            Call pbSub60_CtrlsEnblDsblDictStartUp(fForM)
            '-----------------------------------
            '-------------------------------------------------------------

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Retorna de [ pbSub61_cCtrlsEnblDsblDictBuild ] ap�s" & vbCr & "avaliar todos os controles" & vbCr & " " & vbCr & " "
If gBbDepurandoLv01a Then Stop
            
SkipTo5:

If gBbDepurandoLv01a Then MsgBox "----- SysLoad01_SysDictsLoad ---------------------------------------------------" & vbCr & vbCr & "Fecha o form [ " & sForM & " ] ap�s" & vbCr & "carregar seus dicion�rios"
If gBbDepurandoLv01a Then Stop
            DoCmd.Close acForm, sForM, acSaveNo
SkipTo6:
        
        '-----------------------------------
        '-------------------------------------------------------------
        'encerra a montagem de dicion�rios do Formul�rio ora avaliado
        
        Else
            'Fecha o Form [ SysStart ] pra evitar problemas de exibi��o na carga dos Forms do sistema
            DoCmd.Close acForm, sForM, acSaveNo

        End If
'    'Teste de acesso aos valores armezanados
'    'Set dDict = dictTrggCtrlsInForm(sForM)
'
'    For Each vKey In dictTrggCtrlsInForm(sForM)
'Stop
'        Set clObjFilGrpsByForm = dictTrggCtrlsInForm(sForM)(vKey)
'
'    Next vKey


SKIP_ALL:       'apagar quando sistema estiver finalizado
'Stop
        
    Next oFrmObjct
    '-------------------------------------------------------------
    '-----------------------------------

'If gBbDepurandoLv01b Then MsgBox "teste - carregar form padr�o ------------------------------------------------" & vbCr & ""
'Stop
    
End Sub


Sub FormStatusBar01_Bld(sForM As String, sWarnID As String, sLoadLogWarn As String, Optional sCtrL As String)
'Sub FormStatusBar01_Bld(sForM As String, sLoadLogWarn As String, Optional sCtrL As String, Optional sRecCntCtrl As String, Optional sNoTrgtCtrlsWarn As String, Optional bHLclrWarn As Boolean, Optional sCtrL As String)
'Sub FormStatusBar01_Bld(sForM As String, Optional sNoTrgtCtrlsWarn As String, Optional bHLclrWarn As Boolean, Optional sCtrL As String)
    Dim vA, vB, vC, vD, vE
'Stop

    'Carrega o dict [ dictFormsParams ] com informa��es sobre
    ' falhas na carga de par�metros de controles do [ Form ]
    If Not IsObject(dictFormsParams) Then Set dictFormsParams = New Dictionary
    
    'vA = dictFormsParams.Exists(sForM)
    'vB = IsObject(dictFormsParams(sForM))
    'vC = IsEmpty(dictFormsParams(sForM))
    
    If dictFormsParams.Exists(sForM) Then
        Set clObjFormsParams = dictFormsParams(sForM)
    
    Else
        Set clObjFormsParams = New cls_09aFormsParams
            dictFormsParams.Add sForM, clObjFormsParams
                
    End If
    
    '---------------------------------------------------------
    'Inicia a inclus�o de alertas no [ dictFormsParams ]
    
    'Indica que h� alertas a serem exibidos no [ Form ]
    clObjFormsParams.bForm_ShowWarns = True
    
'Stop
    'Guarda no dict [ clObjFormsParams.dForm_StatusBarWarns ] os alertas que dever�o ser exibidos ao ser carregado o formul�rio
    
    'vA = IsObject(clObjFormsParams.dForm_StatusBarWarns)
    'vB = clObjFormsParams.dForm_StatusBarWarns Is Nothing
    
    
    If clObjFormsParams.dForm_StatusBarWarns.Exists(sWarnID) = True Then
        Set clObjStatusBarWarn = clObjFormsParams.dForm_StatusBarWarns(sWarnID)
        'Set clObjFormsParams.dForm_StatusBarWarns(sWarnID).clObjStatusBarWarn = clObjFormsParams.dForm_StatusBarWarns(sWarnID)
        'clObjFormsParams.dForm_StatusBarWarns.Item(sWarnID) = sLoadLogWarn
        
        
    Else
        Set clObjStatusBarWarn = New cls_09bFormsParamsLoadLogTxt
        clObjFormsParams.dForm_StatusBarWarns.Add sWarnID, clObjStatusBarWarn
        'clObjFormsParams.dForm_StatusBarWarns.Add sWarnID, sLoadLogWarn
        
        'clObjFormsParams.dForm_StatusBarWarns(sWarnID).clObjStatusBarWarn.
        
        'Set clObjStatusBarWarn = clObjFormsParams.dForm_StatusBarWarns(sWarnID)
        
    End If

    clObjStatusBarWarn.sWarnText = sLoadLogWarn
    
    If Not clObjStatusBarWarn.dForm_StatusBarCtrls.Exists(sCtrL) Then
'Stop
        clObjStatusBarWarn.dForm_StatusBarCtrls.Add sCtrL, sCtrL
        
    End If

'    If clObjFormsParams.dForm_StatusBarWarns.Exists(sWarnID) = True Then
'        clObjFormsParams.dForm_StatusBarWarns.Item(sWarnID) = sLoadLogWarn
'
'
'    Else
'        clObjFormsParams.dForm_StatusBarWarns.Add sWarnID, sLoadLogWarn
'
'
'    End If

End Sub


Sub FormStatusBar02_SetWarn(fForM As Form, bShowWarns As Boolean) ', oOjcT As Object) 'fForm As Form, sStatusTxt As String, sStatusTipText As String)
    
    Dim vA, vB
    Dim sForM As String
    Dim cStatusBar As Control

'MsgBox "----- FormStatusBar02_SetWarn --------------------------------------------------" & vbCr & vbCr & "Configura o label [ StatusBar ] pra indicar que h� alertas de carga do sistema." & vbCr & " " & vbCr & " "
'If gBbDepurandoLv02a Then Stop

'Stop
    'Configura o label [ StatusBar ] pra indicar que h� alertas de carga do sistema
    sForM = fForM.Name
    Set cStatusBar = Forms(sForM).Controls("lblStatusBar")
    
    If bShowWarns Then
        cStatusBar.Caption = "H� alertas de carga do sistema: duplo clique AQUI pra exibir."
        cStatusBar.BackColor = GbLngSTATUSbarAlert
    
    Else
        cStatusBar.Caption = "Carga do sistema: SEM OCORR�NCIAS"
        cStatusBar.BackColor = GbLngSTATUSbarNoAlert
    
    End If
    
    
End Sub

Sub FormStatusBar04_OpnLogForm(fForM As Form) ', oOjcT As Object) 'fForm As Form, sStatusTxt As String, sStatusTipText As String)
    Dim vA, vB, vC, vD, vE
    Dim vFormCoords(1) As Variant
    Dim sForM As String
    Dim cStatusBar As Control
    Dim vKey As Variant
    Dim vWarnID As Variant, vCtrlsInWarn As Variant
    Dim sSysLoadingLog As String, sSysLoadingTmp As String, sLogTitle As String, sLogItems As String
    'Dim lngStatusBarClr As Long
    'Dim lngCounT As Long
    'Dim iInT As Integer
    

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Open Log form"
'Stop
    sForM = fForM.Name
        
    'Confirma se h� alertas de carga do sistema pra serem exibidos
    ' do contr�rio n�o abre o Formul�rio de Alertas
    If FormLoad06b_BackFromFormLoad(sForM) Then
'Stop
        'Monta o Texto ser exibido no Log com textos de alerta da carga do sistema
        If IsObject(dictFormsParams) Then
            If dictFormsParams.Exists(sForM) Then
                Set clObjFormsParams = dictFormsParams(sForM)
                        
                For Each vWarnID In clObjFormsParams.dForm_StatusBarWarns
'Stop
                    
                    sLogItems = ""
                    sSysLoadingTmp = ""
                    
                    Set clObjStatusBarWarn = clObjFormsParams.dForm_StatusBarWarns(vWarnID)
                    
                        sLogTitle = clObjStatusBarWarn.sWarnText
                        If gBbDebugOn Then Debug.Print sLogTitle
                    
                        'vB = ""
                        'vC = ""
                        
                        For Each vCtrlsInWarn In clObjStatusBarWarn.dForm_StatusBarCtrls
'Stop
                            vA = IIf(vCtrlsInWarn <> "", "  . ", "") & vCtrlsInWarn
                            sLogItems = IIf(sLogItems <> "", sLogItems & vbCrLf & vA, vA)
                            
                        Next vCtrlsInWarn
'Stop
                        
'                    if gBbDebugOn then       Debug.Print sLogItems
                    vA = IIf(sLogItems <> "", vbCrLf, "")
                    sSysLoadingTmp = sLogTitle & vA & sLogItems

'                    if gBbDebugOn then       Debug.Print sSysLoadingTmp
                    
                    'vA = clObjFormsParams.dForm_StatusBarWarns(vWarnID)
                    'vA = IIf(sSysLoadingLog  <> "", " / ", "")
                    'sSysLoadingLog  = sSysLoadingLog  & vA & clObjFormsParams.dForm_StatusBarWarns(vKey)
                    
                    vA = IIf(sSysLoadingLog <> "", vbNewLine & vbNewLine, "")
                    sSysLoadingLog = sSysLoadingLog & vA & sSysLoadingTmp
                    'sSysLoadingLog = IIf(sSysLoadingLog <> "", sSysLoadingLog & vbNewLine, "") & IIf(vA <> "", "- " & vA, "")
'                    Debug.Print sSysLoadingLog

                Next vWarnID
                
                'Debug.Print sSysLoadingLog
                'sSysLoadingLog = ""
                        
            End If
'Stop
        
        
        End If
        
        'Atribui��o dos valores a [ clObjFormOpenParams ] usados nas rotinas de abertura de formul�rios
        Set clObjFormOpenParams = New cls_09cParamsToOpenForms
        
        With clObjFormOpenParams
               
            '.bIsTrggForm = False
            .bSetPosition = True
            '.fTrgtForm   este valor deve ser definido ap�s a abertura do formul�rio em FormLoad05_OpenForm
            .lngFormLeft = 16140
            .lngFormTop = 1545
            '.lngRecID = 30
            .sFormMode = "StartView"
            .sTrggForM = sForM
            .sTrgtForm = "frm_00(1)cSysLoadLog"
            .vOpenArgs = sSysLoadingLog
        End With
'Stop
        'Inicia as rotinas pra abertura de formul�rio
        ' passa os par�metros do form [ Refer�ncia ]
        Call FormLoad01_FormLoadingStart(clObjFormOpenParams)

        'DoCmd.OpenForm "frm_00(1)cSysLoadLog", , , , , , sForm & "_/_" & sSysLoadingLog
        
        'Diferen�a da esquerda do [ frm_00(1)cSysLoadLog ] para o [ frm_01(1)cProdEstoque ]
        'valores positivos para posicionar o [ fForm ] mais a esquerda
        'vFormCoords(0) = 16140
        'Diferen�a do topo do [ frm_00(1)cSysLoadLog ] para o [ frm_01(1)cProdEstoque ]
        'valores positivos para posicionar o [ fForm ] mais abaixo
        'vFormCoords(1) = 1545
        'Para verificar a posi��o atual dos [ fForms ] na tela execute: call CheckFormPosition na janela de verifica��o imediata
        
        'Call FormLoad08_SetFormPosition("frm_00(1)cSysLoadLog", vFormCoords, "frm_01(1)cProdEstoque", , True)
        
    End If
    
    'DoEvents
'Stop
End Sub
