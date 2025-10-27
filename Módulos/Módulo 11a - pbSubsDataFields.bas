Attribute VB_Name = "Módulo 11a - pbSubsDataFields"

Option Compare Database
Option Explicit

Public Function BuildFilterSQL(Optional sQry As String, Optional sWhere As String, Optional cLstBox As ListBox, Optional bManterWhere As Boolean) As String
Dim vA, vB, vC, sFirstSecSQL As String, sLastSecSQL As String, sSQL As String, iListIndex As Integer, iQryID As Integer
Dim rsDef As Recordset, sQryIDfield As String, sOldWhere As String

    
    If Not cLstBox Is Nothing Then
        iListIndex = cLstBox.ListIndex
        
        If iListIndex > -1 Then
            iQryID = cLstBox.Column(0, iListIndex)
        End If
        
        If sQry = "" Then sSQL = Replace(cLstBox.RowSource, ";", "") Else sSQL = Replace(sQry, ";", "")
        If InStr(sSQL, "SELECT") > 0 Then sSQL = sSQL Else sSQL = Replace(CurrentDb.QueryDefs(sSQL).sql, ";", "")
        If InStr(sSQL, "GROUP BY") > 0 Then
            sFirstSecSQL = Split(sSQL, "GROUP BY")(0)
            sLastSecSQL = "GROUP BY" & Split(sSQL, "GROUP BY")(1)
        ElseIf InStr(sSQL, "ORDER BY") > 0 Then
            sFirstSecSQL = Split(sSQL, "ORDER BY")(0)
            sLastSecSQL = "ORDER BY" & Split(sSQL, "ORDER BY")(1)
        Else
            sFirstSecSQL = sSQL
        End If
        
        Set rsDef = CurrentDb.OpenRecordset(sSQL)
        sQryIDfield = rsDef.Fields(0).Name
        
        If sWhere = "" Then
            sWhere = "WHERE [" & sQryIDfield & "]" & " = " & iQryID
        Else
            sWhere = "WHERE " & Replace(sWhere, "WHERE", "") & " AND [" & sQryIDfield & "]" & " = " & iQryID
        End If
        
    Else
        If InStr(sQry, "SELECT") > 0 Then sSQL = Replace(sQry, ";", "") Else sSQL = Replace(CurrentDb.QueryDefs(sQry).sql, ";", "")
        
        If InStr(sSQL, "GROUP BY") > 0 Then
            sFirstSecSQL = Split(sSQL, "GROUP BY")(0)
            sLastSecSQL = "GROUP BY" & Split(sSQL, "GROUP BY")(1)
        ElseIf InStr(sSQL, "ORDER BY") > 0 Then
            sFirstSecSQL = Split(sSQL, "ORDER BY")(0)
            sLastSecSQL = "ORDER BY" & Split(sSQL, "ORDER BY")(1)
        Else
            sFirstSecSQL = sSQL
        End If
    End If
           
    If InStr(sFirstSecSQL, "WHERE") > 0 Then
        If bManterWhere Then
            If sWhere <> "" Then
                sWhere = sWhere & " AND " & Split(sFirstSecSQL, "WHERE")(1)
            Else
                sWhere = Split(sFirstSecSQL, "WHERE")(1)
            End If
        End If
        sFirstSecSQL = Split(sFirstSecSQL, "WHERE")(0)
        
    End If
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
    Dim bMskdCtrl As Boolean
    Dim sCustomFormat As String
    
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
    
    sQryListBox = BuildFilterSQL(, , cListBox)
    
    Set rsTbE = CurrentDb.OpenRecordset(sQryListBox, dbOpenDynaset, dbReadOnly)
    
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
        
        'Verifica se o controle possui alguma máscara
        If IsObject(dictCtrlBehvrParams(sForM)) Then
        'dictCtrlBehvrParams(sForM)(sCtrL)
            If dictCtrlBehvrParams(sForM).Exists(sDataFieldCtrl) = True Then
                Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sDataFieldCtrl)
                bMskdCtrl = clObjCtrlBehvrParams.bMskdCtrl
            End If
        End If
                    
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
                    
                    'Caso o controle possua máscara, aplica a formatação através da rotina [ MskdTxtbox_TextMask ]
                    If bMskdCtrl Then
                        sCustomFormat = rstFieldDataField.Properties("Format")
                        If InStr(sCustomFormat, ";") > 0 Then sCustomFormat = Split(sCustomFormat, ";")(0)
                        sCustomFormat = Replace(sCustomFormat, """", "")
                        Call MskdTxtbox_TextMask(cDataFieldCtrl, sCustomFormat)
                    End If
                    
                Else
                    
                    'Confirma se o controle ? uma combobox
                    If cDataFieldCtrl.ControlType = acComboBox Or cDataFieldCtrl.ControlType = acListBox Then
                        
                        Erase vSrchItemsCmb
                        
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
                        Erase vSrchItemsCmb
                        'Percorre cada valor de [ vDefItemsCmb ] para verificar se esse valor ? atribu?do ao item selecionado na lista
                        For iItem = 0 To UBound(vDefItemsCmb)
                
                            'Monta o WHERE da consulta
                            sWhere = ""
                            If Not IsNull(vDefItemsCmb(iItem)) Then
                                sWhere = "([" & clObjCtrlDataFieds.sDataField & "]" & " = " & vDefItemsCmb(iItem) & ")"
                                
                                sQuery = BuildFilterSQL(cListBox.RowSource, sWhere, , True)
        
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
                            End If
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
                        'Seleciona o primeiro item
                        cDataFieldCtrl.Value = cDataFieldCtrl.ItemData(0)
                        
                        'Abre o banco pra inicar a busca do registro
                        Set rsTbE = CurrentDb.OpenRecordset(sQryListBox, dbOpenDynaset, dbReadOnly)
                                                
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
    
End Sub

Public Sub PbSubDataFields_Rec(cBtnSaveRec As Control)

    Dim sForM As String, sBtnSaveRec As String, sFilGrp As String, sRecQry As String, sActType As String
    Dim sDtFldRec As String, sDtFldID As String, sDataFieldCtrl As String, sTrgtCtrl As String
    Dim DtFld As Field, DtFldRec As Field, DtFldID As Field
    Dim cLstBox As Control, cDataFieldCtrl As Control
    Dim vA, vB, vC, vKeyTrgtCtrl, vDataFieldCtrl
    Dim rsRecQry As Recordset
    Dim bBoL As Boolean
    
    sBtnSaveRec = cBtnSaveRec.Name
    sForM = cBtnSaveRec.Parent.Name
    
    If Not dictFormCommButtons(sForM).Exists(sBtnSaveRec) Then Exit Sub
    
    sFilGrp = clObjCommButtons.sFilGrp
    sRecQry = clObjCommButtons.sRecQry
    sActType = clObjCommButtons.sActType
    
    For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(sForM)(sFilGrp)
        
        sTrgtCtrl = vKeyTrgtCtrl
        Set clObjCommButtons = dictFormCommButtons(sForM)(sBtnSaveRec)
        Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)(sTrgtCtrl)
    
        Set cLstBox = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)

        If sActType = "SaveEdit" Then
        
            sRecQry = BuildFilterSQL(sRecQry, , cLstBox)
            
            'If InStr(sRecQry, "DISTINCT") > 0 Then sRecQry = Replace(sRecQry, "DISTINCT", "")
            If Not rsRecQry Is Nothing Then rsRecQry.Close: Set rsRecQry = Nothing
            
            'Abre o recordSet da consulta que ser? usada para edição
            Set rsRecQry = CurrentDb.OpenRecordset(sRecQry, dbOpenDynaset)
            
            If rsRecQry.RecordCount = 0 Then GoTo NextTrgt
            'Percorre os [ DataFieldCtrls ] do [ dictFormDataFlds01Grps(sForM)(sFilGrp) ]
            For Each vDataFieldCtrl In dictFormDataFlds01Grps(sForM)(sFilGrp)
                Set clObjCtrlDataFieds = dictFormDataFlds01Grps(sForM)(sFilGrp)(vDataFieldCtrl)
                sDataFieldCtrl = vDataFieldCtrl
                Set cDataFieldCtrl = Forms(sForM).Controls(sDataFieldCtrl)
                
                sDtFldRec = Replace(clObjCtrlDataFieds.sDataField, "IDfk", "ID")

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

            Next vDataFieldCtrl
            
        ElseIf sActType = "SaveNew" Then
            
            'Abre o recordSet da consulta que ser? usada para edição
            Set rsRecQry = CurrentDb.OpenRecordset(sRecQry, dbOpenDynaset)
            
            'Define [ rsRecQry ] para edição de registro
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
    
    Dim sBtnExcRec As String
    Dim sForM As String
    Dim sFilGrp As String
    Dim sRecQry As String
    Dim sActType As String
    Dim vKeyTrgtCtrl As Variant
    Dim sTrgtCtrl As String
    Dim cLstBox As Control
    Dim sSQL As String
    Dim rsTbE As Recordset
    
    sBtnExcRec = cBtnExcRec.Name
    sForM = cBtnExcRec.Parent.Name
    
    If Not dictFormCommButtons(sForM).Exists(sBtnExcRec) Then Exit Sub
    
    sFilGrp = clObjCommButtons.sFilGrp
    sRecQry = clObjCommButtons.sRecQry
    sActType = clObjCommButtons.sActType

    For Each vKeyTrgtCtrl In dictFormFilterGrpTrgts(sForM)(sFilGrp)
        
        sTrgtCtrl = vKeyTrgtCtrl

        Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(sFilGrp)(sTrgtCtrl)
        
        Set cLstBox = Forms(sForM).Controls(clObjTargtCtrlParam.sTargtCtrlName)

        sSQL = BuildFilterSQL(sRecQry, , cLstBox)
        
        Set rsTbE = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)
        
        If rsTbE.RecordCount = 1 Then rsTbE.Delete

        rsTbE.Close
        Set rsTbE = Nothing
        cLstBox.Requery
NextTrgt:
    
    Next vKeyTrgtCtrl
    
End Sub
