Attribute VB_Name = "Módulo 11a - pbSubsDataFields"

Option Compare Database
Option Explicit

Public Function BuildFilterSQL(sQry As String, Optional sWhere As String) As String
    Dim vA, vB, vC: Dim sFirstSecSQL As String: Dim sLastSecSQL As String: Dim sSQL As String
    
    If InStr(sQry, "SELECT") > 0 Then sSQL = sQry Else sSQL = Replace(CurrentDb.QueryDefs(sQry).sql, ";", "")
    
    If InStr(sSQL, "GROUP BY") > 0 Then
        sFirstSecSQL = Split(sSQL, "GROUP BY")(0)
        sLastSecSQL = "GROUP BY" & Split(sSQL, "GROUP BY")(1)
    ElseIf InStr(sSQL, "ORDER BY") > 0 Then
        sFirstSecSQL = Split(sSQL, "ORDER BY")(0)
        sLastSecSQL = "ORDER BY" & Split(sSQL, "ORDER BY")(1)
    End If

    If InStr(sFirstSecSQL, "WHERE") > 0 Then sFirstSecSQL = Split(sFirstSecSQL, "WHERE")(0)
    If InStr(sWhere, "WHERE") = 0 Then sWhere = "WHERE " & sWhere
    sSQL = sFirstSecSQL & sWhere & sLastSecSQL
    
    'Formata a estrutura de [ sSQL ] para facilitar a leitura do usuário
    sSQL = Replace(Replace(Replace(Replace(Replace(Replace(sSQL, vbCrLf, ""), _
                                                    "FROM", vbCrLf & "FROM"), _
                                                    "WHERE", vbCrLf & "WHERE"), _
                                                    "GROUP", vbCrLf & "GROUP"), _
                                                    "ORDER", vbCrLf & "ORDER"), _
                                                    "HAVING", vbCrLf & "HAVING")
    
    BuildFilterSQL = sSQL

End Function


Sub PbSubDataFields_FillFromListbox(cListBox As Control)
    'Congela tela duranta a execução da rotina
    LockWindowUpdate Application.hWndAccessApp
    
    Dim vA, vB, vC
    Dim sQuery As String
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
    Dim sQryOrder As String
    Dim sQryListBox As String
    Dim fForM As Form
    Dim vKeyQDef As Variant
    Dim vKeyField As Variant
    Dim fField As Field
    Dim rsJct As Recordset
    Dim sFieldID As String
    Dim vKeyDataFieldCtrl As Variant
    Dim vWdthsCol As Variant
    Dim vDefItemsCmb() As Variant
    Dim vSrchItemsCmb() As Variant
    Dim qDefJct As QueryDef
    Dim rsTbE As Recordset
    Dim rsDefQry As Recordset
    Dim rsTbECmb As Recordset
    Dim rstFieldDataField As Field
    Dim iListIndex As Integer
    Dim iQryID As Integer
    Dim iItem As Integer
    Dim iConT As Integer
    Dim iColIDCmb As Integer
    Dim sValue As String
    Dim vCamposNecessarios As Variant
    Dim sWhere As String
    
    Set fForM = cListBox.Parent
    sForM = fForM.Name

    '--------------------------------------------------------------------
    '--                                                        ----------
    '--  Fun??o para preencher os campos alvos de uma listbox  ----------
    '--                                                        ----------
    '--------------------------------------------------------------------
    
    '----------------------------  Configura??es necess?rias para funcionamento ------------------------------
    '---------------------------------------------------------------------------------------------------------
    
    '1. Declarar o dicion?rio no m?dulo de vari?veis
    '---------------------------------------------------------------------------------------------------------
    '  'dict para guardar as consultas padr?o dos controles, tanto [ TrgtCtrls ] que j? est?o no dict de [ Targets ]
    '    como tamb?m as Combos e demais Listboxes do [ Form ]
    '  Public dictFormQrysCtrls As New Dictionary
    
    '---------------------------------------------------------------------------------------------------------
    '2. Adicionar a sub CleanDicts para remover os itens do dicion?rio
    '---------------------------------------------------------------------------------------------------------
    'dictFormQrysCtrls.RemoveAll
    '---------------------------------------------------------------------------------------------------------
           
    '3. Adicionar o c?digo abaixo na sub de inicializa??o dos dicion?rios,
    '     teste feito adicionando na sub [ pbSub30_TriggCtrlDictStartUp ]
    '     ap?s a linha [ Case acCheckBox, acOptionGroup, acTextBox, acListBox, acComboBox ]
    '
    '    'C?digo para carregar o dicion?rio com as consultas dos controles do tipo [ acListBox ] e [ acComboBox ]
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
    'Inicia a consulta pra exibi??o dos dados
    '-------------------------------------------
    
    'Identifica o registro selecionado na Listbox
    iListIndex = cListBox.ListIndex
    
    If iListIndex = -1 Then
        If IsObject(dictFormDataFlds01Grps(sForM)(sFilGrp)) Then
            For Each vKeyDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
                Set cDataFieldCtrl = fForM.Controls(vKeyDataFieldCtrl)
                cDataFieldCtrl.Value = ""
                If cDataFieldCtrl.ControlType = acListBox Then cDataFieldCtrl.RowSource = ""
            Next vKeyDataFieldCtrl
        End If
        Exit Sub
    End If
    'Identifica o [ ID ] do registro selecionado, na Tabela da dados
    iQryID = cListBox.Column(0, iListIndex)
    
    'Recupera o SQL da lista
    ' caso a propriedade [ RowSource ] da lista n?o contenha "SELECT" indica que se trata de um nome de consulta
    ' caso contr?rio, indica que j? se trata de um SQL
    If dictFormQrysCtrls(sForM).Exists(cListBox.Name) Then
        sQryListBox = dictFormQrysCtrls(sForM)(cListBox.Name)
    Else
        sQryListBox = cListBox.RowSource
    End If
    
    If InStr(sQryListBox, "SELECT") = 0 Then
        Set qDef = CurrentDb.QueryDefs(sQryListBox)
        sQryListBox = Replace(qDef.sql, ";", "")
    End If
    
    'Abre o banco pra inicar a busca do registro
    Set rsTbE = CurrentDb.OpenRecordset(sQryListBox, dbOpenDynaset, dbReadOnly)
    
    'O nome do campo da Consulta que armazena o ID do registro
    ' ? recuperado pra ser usado na montagem da filtragem, a partir da 1a Coluna da Tabela de Dados
    sQryIDfield = rsTbE.Fields(0).Name
    
    'Montagem dos par?metros de busca
    vA = "[" & sQryIDfield & "]" & " = " & iQryID
    rsTbE.Filter = vA
    Set rsTbE = rsTbE.OpenRecordset

    '------------------------------------------------
    'Exibi??o dos dados recuperados da consulta
    '------------------------------------------------
    
    'Se o [ sFilGrp ] n?o estiver no [ dictFormDataFlds01Grps(sForM) ],
    ' sai da rotina pois n?o h? [ DataFields ] associados ao brupo a serem preenchidos
    If Not IsObject(dictFormDataFlds01Grps(sForM)(sFilGrp)) Then Exit Sub

    'Varre os controles [ DataField ] associados ao [ grupo de filtragem ]
    For Each vKeyDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
        'Sai da rotina caso o ?ltimo registro seja vazio, bug que costuma acontecer no VBA
        If IsEmpty(vKeyDataFieldCtrl) Then Exit Sub
        
        'Define o [ clObjCtrlDataFieds ] referente ao controle ora analisado
        Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vKeyDataFieldCtrl)
        sDataFieldCtrl = vKeyDataFieldCtrl
    
        'Confirma se o controle [ vKeyDataFieldCtrl ] de fato existe no [ Form ]
        If ControlExists(sDataFieldCtrl, fForM) Then
            Set cDataFieldCtrl = fForM.Controls(sDataFieldCtrl)
            
            Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)(cListBox.Name)
            
            'Confirma se [ clObjCtrlDataFieds.sDataField ] ? um dos campos da consulta de [ cListBox.Name ]
            If clObjTargtCtrlParam.dictTrgtQryFields.Exists(clObjCtrlDataFieds.sDataField) Then
                
                vA = clObjTargtCtrlParam.dictTrgtQryFields(clObjCtrlDataFieds.sDataField)
                
                'Verifica se o campo est? no grid da consulta
                If vA = "Grid" And cDataFieldCtrl.ControlType <> acListBox Then
                    'Atribui ? vari?vel tipo Field [ rstFieldDataField ] o campo da consulta da Listbox [ cListBox ],
                    ' indicado em [ clObjCtrlDataFieds.sDataField ] recuperado da TAG do controle [ vKeyDataFieldCtrl ] ora analisado
                    ' e retorna o valor armazenado na tabela de dados
                    Set rstFieldDataField = rsTbE.Fields(clObjCtrlDataFieds.sDataField)
                    
                    'Exibe no controle ora analisado, o valor recuperado na tabela de dados
                    cDataFieldCtrl.Value = rstFieldDataField

                Else
                    
                    'Confirma se o controle ? uma combobox
                    If cDataFieldCtrl.ControlType = acComboBox Or cDataFieldCtrl.ControlType = acListBox Then
                        
                        'Descobre qual a coluna do controle cont?m os dados a serem pesquisados
                        ' para isso, verifica os [ Widths ] das colunas e atribui a [ iColIDCmb ] o n?mero da coluna que possui width ZERO
                        If cDataFieldCtrl.ControlType = acListBox Then
                            iColIDCmb = 0
                        Else
                            vWdthsCol = Split(cDataFieldCtrl.ColumnWidths, ";")
                            For iConT = 0 To UBound(vWdthsCol)
                                If vWdthsCol(iConT) = "0" Then iColIDCmb = iConT
                            Next iConT
                        End If
                        'Recupera o SQL da consulta que alimenta o controle no [ dictFormQrysCtrls(sForm)(cCtrl) ]
                        If InStr(dictFormQrysCtrls(sForM)(sDataFieldCtrl), "SELECT") = 0 Then '4
                            Set qDef = CurrentDb.QueryDefs(dictFormQrysCtrls(sForM)(sDataFieldCtrl))
                            sDefQuerY = qDef.sql
                        Else
                            sDefQuerY = cDataFieldCtrl.RowSource
                        End If
                
                        'Abre o recordset da consulta para capturar os valores que s?o exibidos por padr?o em [ cDataFieldCtrl ]
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
                        
                        'vari?vel usada para redimensionar [ vDefItemsCmb ]
                        iConT = 0
                        
                        bBoL = False
                        'Percorre cada valor de [ vDefItemsCmb ] para verificar se esse valor ? atribu?do ao item selecionado na lista
                        For iItem = 0 To UBound(vDefItemsCmb)
                
                            'Monta o WHERE da consulta
                            If Not IsNull(vDefItemsCmb(iItem)) Then
                                sWhere = "([" & clObjCtrlDataFieds.sDataField & "]" & " = " & vDefItemsCmb(iItem) & ") AND ([" & sQryIDfield & "]" & " = " & iQryID & ")"
                            Else
                                sWhere = "([" & sQryIDfield & "]" & " = " & iQryID & ")"
                            End If
                           
                            sQuery = BuildFilterSQL(sQryListBox, sWhere)

                            'Abre um RecordSet com o filtro
                            Set rsTbECmb = CurrentDb.OpenRecordset(sQuery, dbOpenDynaset, dbReadOnly)
                            
                            'Caso o [ rsTbECmb ] retorne algo, indica que o [ iItem ] est? atribu?do
                            ' ent?o, armazena o [ iItem ] em [ vDefItemsCmb ]
                            If rsTbECmb.RecordCount > 0 Then
                                ReDim Preserve vSrchItemsCmb(iConT)
                                vSrchItemsCmb(iConT) = vDefItemsCmb(iItem)
                                iConT = iConT + 1
                                bBoL = True
                            End If
                            
                            'Fecha o RecordSet
                            rsTbECmb.Close
                            
                        Next iItem
    
                        sFilterCmb = ""
                        'Percorre [ vDefItemsCmb ] para buscar quais itens dever?o ser inclusos em [ cDataFieldCtrl ]
                        If bBoL Then
                            For iItem = 0 To UBound(vSrchItemsCmb)
                                sFilterCmb = sFilterCmb & "([" & sFieldCmb & "]" & " = " & vSrchItemsCmb(iItem) & ")"
                                'Caso ainda n?o seja o ?ltimo item adiciona [ OR ] ao final para continuar a montagem do filtro
                                If iItem < UBound(vSrchItemsCmb) Then sFilterCmb = sFilterCmb & " OR "
                            Next iItem
                        End If
                        
                        If sFilterCmb = "" Then sFilterCmb = "NÃO ENCONTRADO"
                        
                        sDefQuerY = BuildFilterSQL(sDefQuerY, sFilterCmb)

                        'Atribui a nova [ sDefQuerY ] ao [ cDataFieldCtrl ]
                        cDataFieldCtrl.RowSource = sDefQuerY
'                        'Seleciona o primeiro item
                        cDataFieldCtrl.Value = cDataFieldCtrl.ItemData(0)
                        
                        'Abre o banco pra inicar a busca do registro
                        Set rsTbE = CurrentDb.OpenRecordset(sQryListBox, dbOpenDynaset, dbReadOnly)
                        
                        'O nome do campo da Consulta que armazena o ID do registro
                        ' ? recuperado pra ser usado na montagem da filtragem, a partir da 1a Coluna da Tabela de Dados
                        sQryIDfield = rsTbE.Fields(0).Name
                        
                        'Montagem dos par?metros de busca
                        vA = "[" & sQryIDfield & "]" & " = " & iQryID
                        rsTbE.Filter = vA
                        Set rsTbE = rsTbE.OpenRecordset
                        
                        For Each vB In rsTbE.Fields
                            If InStr(vB.Name, Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")) > 0 Then
                                cDataFieldCtrl.Value = vB
                            End If
                        Next vB
                    
                    End If
                    
                End If
                
            Else
                'Inclui o erro no dict de Logs de Carga do sistema
                vA = "Na TAG dos seguintes DataFields foi indicada uma coluna de dados n?o localizada na consulta fonte do [ TargtCtrl ] associado ao controle."
                vB = vbCrLf & "Esses DataFields n?o exibir?o dados."
                sLoadLogWarn = vA & vB
                
                Call FormStatusBar01_Bld(sForM, "MissingDataFieldQryField", sLoadLogWarn, sDataFieldCtrl)

            End If

        End If
        
        'Esvazia [ vDefItemsCmb ]
        ReDim vDefItemsCmb(0)
        vDefItemsCmb(0) = ""
        
    Next vKeyDataFieldCtrl
    
    'Libera atualização de tela
    LockWindowUpdate 0
End Sub

Public Sub PbSubDataFields_Rec(cBtnSaveRec As Control)
    'Congela tela duranta a execução da rotina
    LockWindowUpdate Application.hWndAccessApp
    
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
    Dim vKeyTrgtCtrl
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
    Dim sTrgtCtrl As String
    Dim fField As Field
    Dim sQryOrder As String
    Dim qLstQryDef As QueryDef
    
    sBtnSaveRec = cBtnSaveRec.Name
    sForM = cBtnSaveRec.Parent.Name
    
    If Not dictFormCommButtons(sForM).Exists(sBtnSaveRec) Then
        'Montar mensagem de erro caso o bot?o n?o esteja no dicion?rio
        Exit Sub
    End If
    
    sFilGrp = clObjCommButtons.sFilGrp
    sRecQry = clObjCommButtons.sRecQry
    sActType = clObjCommButtons.sActType
    
    For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(sForM)(sFilGrp)
        
        sTrgtCtrl = vKeyTrgtCtrl
        Set clObjCommButtons = dictFormCommButtons(sForM)(sBtnSaveRec)
        Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)(sTrgtCtrl)
    
        Set cLstBox = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)

        'Abre o recordSet da consulta que ser? usada para edi??o
        Set rsRecQry = CurrentDb.OpenRecordset(sRecQry, dbOpenDynaset)
        
        If sActType = "SaveEdit" Then
            
            'Identifica o registro selecionado na Listbox
            iListIndex = cLstBox.ListIndex
            
            'Identifica o ID do registro selecionado
            If iListIndex > -1 Then
                iQryID = cLstBox.Column(0, iListIndex)

            Else
             GoTo NextTrgt
            End If
            
            sQuerYLstBox = clObjTargtCtrlParam.sClsLstbxSQL_gMAIN
            
            Set rsTbE = CurrentDb.OpenRecordset(sQuerYLstBox, dbOpenDynaset, dbReadOnly)
            
            sQryIDfield = rsTbE.Fields(0).Name
            sDtFldID = Replace(sQryIDfield, "IDfk", "ID")
            
            rsTbE.Close
            Set rsTbE = Nothing
        
            'Aplica filtro na consulta para retornar apenas o item que deve ser editado
            vA = "[" & sQryIDfield & "]" & " = " & iQryID
            rsRecQry.Filter = vA
            Set rsRecQry = rsRecQry.OpenRecordset
            
            If rsRecQry.RecordCount = 0 Then GoTo NextTrgt
            'Percorre os [ DataFieldCtrls ] do [ dictFormDataFlds01Grps(sForM)(sFilGrp) ]
            For Each vDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
                Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vDataFieldCtrl)
                sDataFieldCtrl = vDataFieldCtrl
                Set cDataFieldCtrl = Forms(sForM).Controls(sDataFieldCtrl)
                
                sDtFldRec = Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")

                    'Define [ rsRecQry ] para edi??o de registro
                    rsRecQry.Edit
                    
                    Set DtFldRec = Nothing
                    
                    For Each DtFld In rsRecQry.Fields
                        If DtFld.Name Like sDtFldRec & "*" Then Set DtFldRec = DtFld
                    Next DtFld
                    
                    'Se o campo foi localizado, altera o valor
                    If Not DtFldRec Is Nothing Then DtFldRec = cDataFieldCtrl.Value
                    
                               Debug.Print sDataFieldCtrl
                    'Salva altera??es
                    rsRecQry.Update

            Next vDataFieldCtrl
            
            
        ElseIf sActType = "SaveNew" Then

            
            'Define [ rsRecQry ] para edi??o de registro
            rsRecQry.AddNew
            
            'Verifica todas as informações que são necessárias para gravação dos dados
            For Each DtFld In rsRecQry.Fields
                bBoL = False
                If DtFld.CollectionIndex > 0 Then
                    Debug.Print DtFld.Name
                    For Each vA In dictFormDataFlds01Grps(sForM)(sFilGrp)
                        Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vA)
                        If DtFld.Name Like clObjCtrlDataFieds.sDataField & "*" Or clObjCtrlDataFieds.sDataField Like DtFld.Name & "*" Then
                            Set cDataFieldCtrl = Forms(sForM).Controls(clObjCtrlDataFieds.sCtrlDataField)
                            If Not cDataFieldCtrl.ControlType = acListBox And cDataFieldCtrl.Value <> "" Then
                                DtFld = cDataFieldCtrl.Value
                                bBoL = True
                                Exit For
                            End If
                        End If
                    Next vA
                    If Not bBoL Then
                        For Each vA In dictFieldsSelectedItem(sForM)(sFilGrp)
                            If Replace(DtFld.Name, "IDfk", "ID") = Replace(vA, "IDfk", "ID") Then
                                DtFld = dictFieldsSelectedItem(sForM)(sFilGrp)(vA)
                                Exit For
                            End If
                        Next vA
                    End If
                End If
                
            Next DtFld
                    
            rsRecQry.Update
            
      End If
NextTrgt:
    Next vKeyTrgtCtrl
    
    'Atualiza os [ TrgtCtrls ] do grupo
    For Each vA In dictFormFilterGrpTrgts(sForM)(sFilGrp)
        Set cLstBox = Forms(sForM).Controls(vA)
        cLstBox.Requery
        Call PbSubDataFields_FillFromListbox(cLstBox)
    Next vA

    LockWindowUpdate 0
End Sub

Public Sub PbSubDataFields_Delete(cBtnExcRec As Control)
    'Congela tela duranta a execução da rotina
    LockWindowUpdate Application.hWndAccessApp
    
    Dim vA, vB, vC
    Dim sBtnExcRec As String
    Dim sForM As String
    Dim sFilGrp As String
    Dim sRecQry As String
    Dim sActType As String
    Dim vKeyTrgtCtrl As Variant
    Dim sTrgtCtrl As String
    Dim cLstBox As Control
    Dim iListIndex As Integer
    Dim iQryID As Integer
    Dim sSQL As String
    Dim qDef As QueryDef
    Dim sWhere As String
    Dim rsTbE As Recordset
    Dim sQryIDfield As String
    Dim bBoL As Boolean
    
    bBoL = False
    
    sBtnExcRec = cBtnExcRec.Name
    sForM = cBtnExcRec.Parent.Name
    
    If Not dictFormCommButtons(sForM).Exists(sBtnExcRec) Then
        'Montar mensagem de erro caso o bot?o n?o esteja no dicion?rio
        Exit Sub
    End If
    
    sFilGrp = clObjCommButtons.sFilGrp
    sRecQry = clObjCommButtons.sRecQry
    sActType = clObjCommButtons.sActType
    For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(sForM)(sFilGrp)
        
        sTrgtCtrl = vKeyTrgtCtrl
        Set clObjCommButtons = dictFormCommButtons(sForM)(sBtnExcRec)
        Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)(sTrgtCtrl)
    
        Set cLstBox = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)
        
        'Identifica o registro selecionado na Listbox
        iListIndex = cLstBox.ListIndex
        
        'Identifica o ID do registro selecionado
        If iListIndex > -1 Then
            iQryID = cLstBox.Column(0, iListIndex)
            bBoL = True
        Else
            GoTo NextTrgt
        End If
        
        Set qDef = CurrentDb.QueryDefs(sRecQry)
        
        sQryIDfield = qDef.Fields(0).Name
        sSQL = qDef.sql

        sSQL = Replace(sSQL, ";", "")
        If InStr(sSQL, "GROUP BY") > 0 Then sSQL = Split(sSQL, "GROUP BY")(0)
        If InStr(sSQL, "ORDER BY") > 0 Then sSQL = Split(sSQL, "ORDER BY")(0)
        
        sWhere = "WHERE [" & sQryIDfield & "]" & " = " & iQryID
        
        sSQL = sSQL & vbCrLf & sWhere
        
        sSQL = BuildFilterSQL(sRecQry, sWhere)
        Set rsTbE = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)
        
        If rsTbE.RecordCount = 1 Then
            rsTbE.Delete
        Else
            'Mensagem de erro caso o registro não seja encontrado ou
            ' haja mais de um registro para o [ iQryID ] recuperado
        End If
    
        rsTbE.Close
        Set rsTbE = Nothing
        cLstBox.Requery
NextTrgt:
    
    Next vKeyTrgtCtrl
    
    If Not bBoL Then MsgBox "É necessário selecionar algum item da lista antes de clicar no botão [ Excluir ]", vbCritical
    'Libera atualização de tela
    LockWindowUpdate 0
End Sub
