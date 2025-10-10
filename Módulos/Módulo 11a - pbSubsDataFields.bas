Attribute VB_Name = "Módulo 11a - pbSubsDataFields"
Option Compare Database
Option Explicit


Sub PbSubFillFieldsByList(cListBox As Control)

    Dim vA, vB, vC
    
    Dim sQuerY As String
    Dim sDefQuerY As String
    Dim sForM As String
    Dim sDataFieldCtrl As String
    Dim cDataFieldCtrl As Control
    Dim sFilGrp As String
    Dim sQryIDfield As String
    Dim sFieldCmb As String
    Dim sFilterCmb As String
    Dim bBoL As Boolean
    Dim qDef As QueryDef
    Dim sSQLtablesString As String
    Dim sLoadLogWarn As String
    
    Dim fForM As Form

    
    Dim vKeyDataFieldCtrl As Variant
    Dim vWdthsCol As Variant
    Dim vDefItemsCmb() As Variant
    Dim vSrchItemsCmb() As Variant
    
    Dim rsTbE As Recordset
    Dim rsDefQry As Recordset
    Dim rsTbECmb As Recordset
    
    Dim rstFieldDataField As Field

    Dim iListIndex As Integer
    Dim iQryID As Integer
    Dim iItem As Integer
    Dim iCont As Integer
    Dim iColIDCmb As Integer

    Set fForM = cListBox.Parent
    sForM = fForM.Name
'Stop

    '--------------------------------------------------------------------
    '--                                                        ----------
    '--  Função para preencher os campos alvos de uma listbox  ----------
    '--                                                        ----------
    '--------------------------------------------------------------------
    
    '----------------------------  Configurações necessárias para funcionamento ------------------------------
    '---------------------------------------------------------------------------------------------------------
    
    '1. Declarar o dicionário no módulo de variáveis
    '---------------------------------------------------------------------------------------------------------
    '  'dict para guardar as consultas padrão dos controles, tanto [ TrgtCtrls ] que já estão no dict de [ Targets ]
    '    como também as Combos e demais Listboxes do [ Form ]
    '  Public dictFormQrysCtrls As New Dictionary
    
    '---------------------------------------------------------------------------------------------------------
    '2. Adicionar a sub CleanDicts para remover os itens do dicionário
    '---------------------------------------------------------------------------------------------------------
    'dictFormQrysCtrls.RemoveAll
    '---------------------------------------------------------------------------------------------------------
           
    '3. Adicionar o código abaixo na sub de inicialização dos dicionários,
    '     teste feito adicionando na sub [ pbSub30_TriggCtrlDictStartUp ]
    '     após a linha [ Case acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox ]
    '
    '    'Código para carregar o dicionário com as consultas dos controles do tipo [ acListBox ] e [ acComboBox ]
    '    '---------------------------------------------------------------------------------------------------------
    '    'If cTriggCtrl.ControlType = acComboBox Or cTriggCtrl.ControlType = acListBox Then
    '    '    If Not IsObject(dictFormQrysCtrls(sForm)) Then Set dictFormQrysCtrls(sForm) = New Dictionary
    '    '    If Not dictFormQrysCtrls(sForm).Exists(cTriggCtrl.Name) Then
    '    '        dictFormQrysCtrls(sForm).Add cTriggCtrl.Name, cTriggCtrl.RowSource
    '    '    End If
    '    'End If
    '    '---------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------

    
    'Recupera o grupo de filtragem do [ ListBox ]
    ' pra depois recuperar os controles [ cDataFieldCtrl ] do [ Grupo ]
    sFilGrp = dictTrgtCtrlsFilterGrps(sForM)(cListBox.Name)
    
    '-------------------------------------------
    'Inicia a consulta pra exibição dos dados
    '-------------------------------------------
    
    'Identifica o registro selecionado na Listbox
    iListIndex = cListBox.ListIndex

    'Identifica o [ ID ] do registro selecionado, na Tabela da dados
    iQryID = cListBox.Column(0, iListIndex)
    
    'Recupera o SQL da lista
    ' caso a propriedade [ RowSource ] da lista não contenha "SELECT" indica que se trata de um nome de consulta
    ' caso contrário, indica que já se trata de um SQL
    If dictFormQrysCtrls(sForM).Exists(cListBox.Name) Then
        sQuerY = dictFormQrysCtrls(sForM)(cListBox.Name)
    Else
        sQuerY = cListBox.RowSource
    End If
    
    If InStr(sQuerY, "SELECT") = 0 Then
        Set qDef = CurrentDb.QueryDefs(sQuerY)
        sQuerY = Replace(qDef.sql, ";", "")
    End If
    
    'Abre o banco pra inicar a busca do registro
    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    
    'O nome do campo da Consulta que armazena o ID do registro
    ' é recuperado pra ser usado na montagem da filtragem, a partir da 1a Coluna da Tabela de Dados
    sQryIDfield = rsTbE.Fields(0).Name
    
    'Montagem dos parâmetros de busca
    vA = "[" & sQryIDfield & "]" & " = " & iQryID
    rsTbE.Filter = vA
    Set rsTbE = rsTbE.OpenRecordset

    '------------------------------------------------
    'Exibição dos dados recuperados da consulta
    '------------------------------------------------
    
    'Se o [ sFilGrp ] não estiver no [ dictFormDataFlds01Grps(sForM) ],
    ' sai da rotina pois não há [ DataFields ] associados ao brupo a serem preenchidos
    If Not IsObject(dictFormDataFlds01Grps(sForM)(sFilGrp)) Then Exit Sub

    'Varre os controles [ DataField ] associados ao [ grupo de filtragem ]
    For Each vKeyDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
        'Sai da rotina caso o último registro seja vazio, bug que costuma acontecer no VBA
        If IsEmpty(vKeyDataFieldCtrl) Then Exit Sub
        
        'Define o [ clObjCtrlDataFieds ] referente ao controle ora analisado
        Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vKeyDataFieldCtrl)
        sDataFieldCtrl = vKeyDataFieldCtrl
    
        'Confirma se o controle [ vKeyDataFieldCtrl ] de fato existe no [ Form ]
        If ControlExists(sDataFieldCtrl, fForM) Then
            Set cDataFieldCtrl = fForM.Controls(sDataFieldCtrl)





    
            
            '----------------------------------------------------------------------------------------------
            '----------------------------------------------------------------------------------------------
            'Checar se o campo [ clObjCtrlDataFieds.sDataField ] indicado no DataField [ sDataFieldCtrl ] existe no grid da consulta do TargtCtrl [ cListBox ]
            ' caso positivo exibe, no controle, o valor referente à coluna de dados armazenado na tabela
            
            ' caso negativo procura em todas as tabelas e Queries usadas na consulta
            ' se mesmo assim não tiver localizado indica o erro no Log de Carga do sistema
            ' se por outro lado encontrar significa que é um campo multiselect e, nesse caso, faz o tratamento
            ' pra recuperar quais valores devem ser exibidos no controle
            NstdVarQryFld = GetFldInQryGrid(sForM, cListBox.Name, clObjCtrlDataFieds.sDataField)
            '-----------------------------------------------------------------------------------
            
'MsgBox "Avalia DataField: [ " & sDataFieldCtrl & " ]"
'Stop
            '----------------------------------------------------------------------------------------------
            If NstdVarQryFld.bFoundQryFld Then
                'Atribui à variável tipo Field [ rstFieldDataField ] o campo da consulta da Listbox [ cListBox ],
                ' indicado em [ clObjCtrlDataFieds.sDataField ] recuperado da TAG do controle [ vKeyDataFieldCtrl ] ora analisado
                ' e retorna o valor armazenado na tabela de dados
                Set rstFieldDataField = rsTbE.Fields(clObjCtrlDataFieds.sDataField)
                
                'Exibe no controle ora analisado, o valor recuperado na tabela de dados
                cDataFieldCtrl.Value = rstFieldDataField
                
            'Se o campo não for encontrado no grid da consulta, procura em todas as tabelas e Queries usadas na consulta
            Else
                '-----------------------------------------------------------------------------------
                'Checar se o campo [ clObjCtrlDataFieds.sDataField ] indicado no DataField [ sDataFieldCtrl ] existe em alguma das tabelas
                ' se não existir carrega no log de carga do sistema
                
                
                If IsObject(dictFormFilterGrps(sForM)) Then
                
                    If dictFormFilterGrps(sForM).Exists(sFilGrp) = True Then
'Stop
                        Set clObjTargtCtrlParam = dictFormFilterGrps(sForM)(sFilGrp)
                    
                        sSQLtablesString = clObjTargtCtrlParam.sClsLstbxSQL_aSELECT & " " & clObjTargtCtrlParam.sClsLstbxSQL_bFROM
                        NstdVarQryFld = GetFldInQryGridTbls(sForM, cListBox.Name, sSQLtablesString, clObjCtrlDataFieds.sDataField)
'Stop
                        vA = NstdVarQryFld.bFoundQryFld
                    End If
                End If
                    
                'Se o campo da consulta não tiver sido encontrado na tabela de dados exibe o alerta
                ' e não inclui o controle no dicionário [ dictTrgg01CtrlsInGrp(sFilGrp) ]
                If Not NstdVarQryFld.bFoundQryFld Then
                    
                    'Inclui o erro no dict de Logs de Carga do sistema
                    vA = "Na TAG dos seguintes DataFields foi indicada uma coluna de dados não localizada na consulta fonte do [ TargtCtrl ] associado ao controle."
                    vB = vbCrLf & "Esses DataFields não exibirão dados."
                    sLoadLogWarn = vA & vB
                    
                    Call FormStatusBar01_Bld(sForM, "MissingDataFieldQryField", sLoadLogWarn, sDataFieldCtrl)
'                    Exit Sub
                
                Else
                
                    'Confirma se o controle é uma combobox
                    If cDataFieldCtrl.ControlType = acComboBox Then
            
                        'Descobre qual a coluna do controle contém os dados a serem pesquisados
                        ' para isso, verifica os [ Widths ] das colunas e atribui a [ iColIDCmb ] o número da coluna que possui width ZERO
                        vWdthsCol = Split(cDataFieldCtrl.ColumnWidths, ";")
                        For iCont = 0 To UBound(vWdthsCol)
                            If vWdthsCol(iCont) = "0" Then iColIDCmb = iCont
                        Next iCont
    
                        'Recupera o SQL da consulta que alimenta o controle no [ dictFormQrysCtrls(sForm)(cCtrl) ]
                        If InStr(dictFormQrysCtrls(sForM)(sDataFieldCtrl), "SELECT") = 0 Then '4
                            Set qDef = CurrentDb.QueryDefs(dictFormQrysCtrls(sForM)(sDataFieldCtrl))
                            sDefQuerY = qDef.sql
                        Else
                            sDefQuerY = cDataFieldCtrl.RowSource
                        End If
                
                        'Abre o recordset da consulta para capturar os valores que são exibidos por padrão em [ cDataFieldCtrl ]
                        Set rsDefQry = CurrentDb.OpenRecordset(sDefQuerY, dbOpenDynaset, dbReadOnly)
                        rsDefQry.MoveLast: rsDefQry.MoveFirst
                        
                        'Captura o nome do campo referente ao item buscado
                        sFieldCmb = rsDefQry.Fields(iColIDCmb).Name
                        
                        'Insere os dados da consulta na [ vDefItemsCmb ]
                        ReDim vDefItemsCmb(rsDefQry.RecordCount - 1)
                        For iItem = 0 To UBound(vDefItemsCmb)
                            vDefItemsCmb(iItem) = rsDefQry.Fields(iColIDCmb)
                            rsDefQry.MoveNext
                        Next iItem
                        
                        rsDefQry.Close
                        
                        'variável usada para redimensionar [ vDefItemsCmb ]
                        iCont = 0
                        
                        bBoL = False
                        'Percorre cada item de [ vDefItemsCmb ] para verificar se o setor é atribuído ao produto selecionado na lista
                        For iItem = 0 To UBound(vDefItemsCmb)
                
                            'Monta o WHERE da consulta
                            vA = "([" & clObjCtrlDataFieds.sDataField & "]" & " = " & vDefItemsCmb(iItem) & ") AND ([" & sQryIDfield & "]" & " = " & iQryID & ")"
                
                            'Modifica consulta da lista filtrando pelo [ sQryIDfield ] e pelo valor do RecordSet do campo
                            If InStr(sQuerY, "WHERE") > 0 Then sQuerY = Split(sQuerY, "WHERE")(0)
                            If InStr(sQuerY, "ORDER BY") > 0 Then
                                vB = "ORDER BY " & Split(sQuerY, "ORDER BY")(1) & ";"
                                sQuerY = Split(sQuerY, "ORDER BY")(0)
                                sQuerY = sQuerY & " WHERE " & vA & vB
                            Else
                                sQuerY = Replace(sQuerY, ";", "") & " WHERE " & vA & ";"
                            End If
                            
                            'Abre um RecordSet com o filtro
                            Set rsTbECmb = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
                            
                            'Caso o [ rsTbECmb ] retorne algum valor, indica que o setor está atribuído
                            ' armazena os setores atribuídos na [ vDefItemsCmb ]
                            If rsTbECmb.RecordCount > 0 Then
                                ReDim Preserve vSrchItemsCmb(iCont)
                                vSrchItemsCmb(iCont) = vDefItemsCmb(iItem)
                                iCont = iCont + 1
                                bBoL = True
                            End If
                            
                            'Fecha o RecordSet
                            rsTbECmb.Close
                            
                        Next iItem
    
                        sFilterCmb = ""
                        'Percorre [ vDefItemsCmb ] para buscar quais itens deverão ser inclusos em [ cDataFieldCtrl ]
                        If bBoL Then
                            For iItem = 0 To UBound(vSrchItemsCmb)
                                sFilterCmb = sFilterCmb & "([" & sFieldCmb & "]" & " = " & vSrchItemsCmb(iItem) & ")"
                                'Caso ainda não seja o último item adiciona [ OR ] ao final para continuar a montagem do filtro
                                If iItem < UBound(vSrchItemsCmb) Then sFilterCmb = sFilterCmb & " OR "
                            Next iItem
                        End If
                        'sFilterCmb = Left(sFilterCmb, Len(sFilterCmb) - 3)
                        
                        'Remonta a [ sDefQuerY ] aplicando o filtro dos items que devem ser exibidos
                        If InStr(sDefQuerY, "WHERE") > 0 Then sDefQuerY = Split(sDefQuerY, "WHERE")(0)
                        If InStr(sDefQuerY, "ORDER BY") > 0 Then '4
                            vB = "ORDER BY " & Split(sDefQuerY, "ORDER BY")(1) & ";"
                            sDefQuerY = Split(sDefQuerY, "ORDER BY")(0)
                            sDefQuerY = sDefQuerY & " WHERE " & sFilterCmb & vB
                            sDefQuerY = Replace(sDefQuerY, ";", "")
                        Else
                            sDefQuerY = Replace(sDefQuerY, ";", "") & " WHERE " & sFilterCmb & ";"
                        End If
                        
                        'Atribui a nova [ sDefQuerY ] ao [ cDataFieldCtrl ]
                        cDataFieldCtrl.RowSource = sDefQuerY
                        'Seleciona o primeiro item
                        cDataFieldCtrl.Value = cDataFieldCtrl.ItemData(0)
                
                    End If
                    
                End If
            
            End If
        
'Stop
            '----------------------------------------------------------------------------------------------
            '----------------------------------------------------------------------------------------------




        End If
        
        'Esvazia [ vDefItemsCmb ]
        ReDim vDefItemsCmb(0)
        vDefItemsCmb(0) = ""
        
    Next vKeyDataFieldCtrl

End Sub

Public Sub PbSubRecDataFields(cBtnSaveRec As Control)

    Dim vA, vB, vC
    
    Dim sForM As String
    Dim sBtnSaveRec As String
    Dim sFilGrp As String
    Dim sRecQry As String
    Dim sActType As String
    Dim iListIndex As Integer
    Dim sQuerYLstBox As String
    Dim iQryID As Integer
    Dim iItemID As Integer
    Dim sQryIDfield As String
    Dim cLstBox As Control
    
    Dim vDataFieldCtrl
    Dim rsTbE As Recordset
    Dim rsRecQry As Recordset
    Dim rsRecQry2 As Recordset
    Dim DtFld As Field
    Dim DtFldRec As Field
    Dim DtFldID As Field
    
    Dim sDtFldRec As String
    Dim sDtFldID As String
    Dim cDataFieldCtrl As Control
    Dim sDataFieldCtrl As String
    Dim qDef As QueryDef
    Dim bBoL As Boolean
    
    sBtnSaveRec = cBtnSaveRec.Name
    sForM = cBtnSaveRec.Parent.Name
    
    If Not dictFormCommButtons(sForM).Exists(sBtnSaveRec) Then
        'Montar mensagem de erro caso o botão não esteja no dicionário
        Exit Sub
    End If
    
    sFilGrp = clObjCommButtons.sFilGrp
    sRecQry = clObjCommButtons.sRecQry
    sActType = clObjCommButtons.sActType
    
    Set clObjCommButtons = dictFormCommButtons(sForM)(sBtnSaveRec)
    Set clObjTargtCtrlParam = dictFormFilterGrps(sForM)(sFilGrp)

    Set cLstBox = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)
    
    'Identifica o registro selecionado na Listbox
    iListIndex = cLstBox.ListIndex
    
    'Identifica o ID do registro selecionado
    If iListIndex > -1 Then iQryID = cLstBox.Column(0, iListIndex)
    
    sQuerYLstBox = clObjTargtCtrlParam.sClsLstbxSQL_eMAIN
    
    Set rsTbE = CurrentDb.OpenRecordset(sQuerYLstBox, dbOpenDynaset, dbReadOnly)
    
    sQryIDfield = rsTbE.Fields(0).Name
    sDtFldID = Replace(sQryIDfield, "IDfk", "ID")
    
    rsTbE.Close
    Set rsTbE = Nothing
    
    'Abre o recordSet da consulta que será usada para edição
    Set rsRecQry = CurrentDb.OpenRecordset(sRecQry, dbOpenDynaset)
        
    If sActType = "SaveEdit" Then
    
        'Aplica filtro na consulta para retornar apenas o item que deve ser editado
        vA = "[" & sQryIDfield & "]" & " = " & iQryID
        rsRecQry.Filter = vA
        Set rsRecQry = rsRecQry.OpenRecordset
        
        'Percorre os [ DataFieldCtrls ] do [ dictFormDataFlds01Grps(sForM)(sFilGrp) ]
        For Each vDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
            Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vDataFieldCtrl)
            sDataFieldCtrl = vDataFieldCtrl
            Set cDataFieldCtrl = Forms(sForM).Controls(sDataFieldCtrl)
            
            sDtFldRec = Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")
            
            'Se houver alguma consulta no parâmetro [ RecQry ] do controle
            If clObjCtrlDataFieds.sRecQry <> "" Then
                If cDataFieldCtrl.ControlType = acComboBox Or cDataFieldCtrl.ControlType = acListBox Then
                    'Abre a consulta indicada pelo parâmetro [ RecQry ]
                    Set rsRecQry2 = CurrentDb.OpenRecordset(clObjCtrlDataFieds.sRecQry, dbOpenDynaset)
                    
                    Set DtFldID = Nothing
                    Set DtFldRec = Nothing
                    
                    For Each DtFld In rsRecQry2.Fields
                        If DtFld.Name Like sDtFldID & "*" Then sDtFldID = DtFld.Name
                        If DtFld.Name Like sDtFldRec & "*" Then Set DtFldRec = DtFld
                    Next DtFld
                    
                    'Aplica filtro na consulta para retornar apenas o item que deve ser editado
                    vA = "[" & sDtFldID & "]" & " = " & iQryID
                    rsRecQry2.Filter = vA
                    Set rsRecQry2 = rsRecQry2.OpenRecordset
                    
                    'Apaga todos os registros referentes ao [ iQryID ]
                    If rsRecQry2.RecordCount > 0 Then
                        Do While Not rsRecQry2.EOF
                            rsRecQry2.Delete
                            rsRecQry2.Update
                            rsRecQry2.MoveNext
                        Loop
                        rsRecQry2.Update
                    End If
    
                    If Not DtFldRec Is Nothing Then
                        For iItemID = 0 To cDataFieldCtrl.ListCount - 1
                            If Not IsNull(cDataFieldCtrl.ItemData(iItemID)) Then
                                rsRecQry2.AddNew
                                rsRecQry2.Fields(sDtFldID) = iQryID
                                DtFldRec = cDataFieldCtrl.ItemData(iItemID)
                                rsRecQry2.Update
                            End If
                        Next iItemID
                    End If
    
                    rsRecQry2.Close
                    Set rsRecQry2 = Nothing
                    
                End If
            'Para controles que não possuem o parâmetro [ RecQry ]
            Else
                'Define [ rsRecQry ] para edição de registro
                rsRecQry.Edit
                
                Set DtFldRec = Nothing
                
                For Each DtFld In rsRecQry.Fields
                    If DtFld.Name Like sDtFldRec & "*" Then Set DtFldRec = DtFld
                Next DtFld
                
                'Se o campo foi localizado, altera o valor
                If Not DtFldRec Is Nothing Then DtFldRec = cDataFieldCtrl.Value
                
                Debug.Print sDataFieldCtrl
                'Salva alterações
                rsRecQry.Update
                
            End If
        Next vDataFieldCtrl
    
    ElseIf sActType = "SaveNew" Then
                
        'Define [ rsRecQry ] para edição de registro
        rsRecQry.AddNew
    
        'Percorre os [ DataFieldCtrls ] do [ dictFormDataFlds01Grps(sForM)(sFilGrp) ]
        For Each vDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
            Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vDataFieldCtrl)
            sDataFieldCtrl = vDataFieldCtrl
            Set cDataFieldCtrl = Forms(sForM).Controls(sDataFieldCtrl)
            
    
            'Verifica se a consulta de gravação está indicada no [ RecQry ] do controle
            If clObjCtrlDataFieds.sRecQry = "" Then
              
                sDtFldRec = Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")
                
                Set DtFldRec = Nothing
                
                For Each DtFld In rsRecQry.Fields
                    If DtFld.Name Like sDtFldRec & "*" Then Set DtFldRec = DtFld
                Next DtFld
                
                'Se o campo foi localizado, altera o valor
                If Not DtFldRec Is Nothing And cDataFieldCtrl.Value <> "" Then DtFldRec = cDataFieldCtrl.Value
                Debug.Print sDataFieldCtrl & " - " & cDataFieldCtrl.Value
            End If
        Next vDataFieldCtrl
        
        'Salva alterações
        rsRecQry.Update
        
        rsRecQry.MoveLast
        iQryID = rsRecQry(sDtFldID)
        
        rsRecQry.Close
        Set rsRecQry = Nothing
        
        For Each vDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
            Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vDataFieldCtrl)
            
            'Verifica se a consulta de gravação está indicada no [ RecQry ] do controle
            If clObjCtrlDataFieds.sRecQry <> "" Then
                sDataFieldCtrl = vDataFieldCtrl
                Set cDataFieldCtrl = Forms(sForM).Controls(sDataFieldCtrl)
                
                sDtFldRec = Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")
                sRecQry = clObjCtrlDataFieds.sRecQry
                
                Set rsRecQry = CurrentDb.OpenRecordset(sRecQry, dbOpenDynaset)
                
                Set DtFldRec = Nothing
                Set DtFldID = Nothing
                
                For Each DtFld In rsRecQry.Fields
                    If DtFld.Name Like sDtFldRec & "*" Then Set DtFldRec = DtFld
                    If DtFld.Name Like sDtFldID & "*" Then Set DtFldID = DtFld
                Next DtFld
                
                If Not DtFldRec Is Nothing And Not DtFldID Is Nothing Then
                    For iItemID = 0 To cDataFieldCtrl.ListCount - 1
                        rsRecQry.AddNew
                        DtFldID = iQryID
                        DtFldRec = cDataFieldCtrl.ItemData(iItemID)
                        rsRecQry.Update
                    Next iItemID
                End If
                
                rsRecQry.Close
                Set rsRecQry = Nothing
                
                Debug.Print sDtFldRec & " - "; bBoL
    
            End If
        Next vDataFieldCtrl
        
    End If
    
    cLstBox.Requery

End Sub
