Attribute VB_Name = "_Testes2"
Option Compare Database
Option Explicit


Function GetSortedSQL(sSQLList As String, sColToOrder As String, sOrderMode As String)
    Dim sFrstSQL As String

    If InStr(sSQLList, "ORDER BY") > 0 Then sFrstSQL = Split(sSQLList, "ORDER BY")(0) Else sFrstSQL = sSQLList
    GetSortedSQL = sFrstSQL & "ORDER BY " & sColToOrder & " " & sOrderMode
    
End Function



'Function GetFldInQryGrid(sForM As String, sTrgtCtrl As String, sQryField As String) As vCheckQryFld
'    Dim vA, vB, vC
'    Dim rsTbE As Recordset
'    Dim sQuerY As String
'    Dim sWhere As String
'    Dim lngFoundRecs As Long
'    Dim fField As Field
'
'    'Abre a consulta que será usada pra filtragem e confirma se o campo de consulta
'    ' informado nos parâmetros do [ TriggCtrl ] existe
'    sQuerY = Forms(sForM).Controls(sTrgtCtrl).RowSource
'    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
'
'    For Each fField In rsTbE.Fields
'        If fField.Name = sQryField Then GetFldInQryGrid.bFoundQryFld = True
'
'    Next fField
'
'    GetFldInQryGrid.sQry = sQuerY
'
'    Set rsTbE = Nothing
'
'End Function
'
'
'Function GetFldInQryGridTbls(sForM As String, sTrgtCtrl As String, sSQLtablesString As String, sQryField As String) As vCheckQryFld
'
'    Dim vA, vB, vC
'    Dim sRegExPattern As String
'
'    Dim rsTbE As Recordset
'    Dim sQuerY As String
'    Dim fField As Field
'
'    Dim regEx As New RegExp
'    Dim mcRegExMatchColl As MatchCollection
'    Dim vRegExKey As Variant
'    Dim vKeyField As Variant
'    Dim vKey As Variant
'
'    Dim db As DAO.Database
'    'Dim tdf As DAO.TableDef
'    'Dim fld As DAO.Field
'
'    '--------------------------------------------------------------------------------------------
'    ' confirma se o campo de consulta [ sQryField ] existe na consulta ou em alguma
'    '  das tabelas da consulta do controle [ sTrgtCtrl ]
'    '--------------------------------------------------------------------------------------------
''Stop
'
'
'    '--------------------------------------------------------------------------------------------
'    'Campo ora avaliado não foi localizado no Grid da consulta,
'    ' continua a pesquisa agora nas [ Tabelas ] usadas na consulta
'    If Not GetFldInQryGridTbls.bFoundQryFld Then
''Stop
'
'    '    'Inicializa o regex object
'    '    Set objRgEXregex = CreateObject("VBScript.RegExp")
''Stop
'
'        regEx.Global = True
'        regEx.IgnoreCase = False
'
'        'Define o padrão a ser buscado
'        sRegExPattern = "\[(tbl)_[0-9]*\([0-9]*\)[a-zA-ZçÇ0-9(\)-_]*\]"
'        regEx.pattern = sRegExPattern
'
'        'vA = regEx.Test(sSQLtablesString )
'
'        Set mcRegExMatchColl = regEx.Execute(sSQLtablesString)
'        'Debug.Print sSQLtablesString
''Stop
'        'Monta dicionário com todas as TABELAS usadas no trecho [ FROM ] da consulta
'        For Each vRegExKey In mcRegExMatchColl
'            vA = vRegExKey.Value
'            If Not dictTempDict.Exists(vRegExKey.Value) Then dictTempDict.Add vRegExKey.Value, vRegExKey
''Stop
'        Next vRegExKey
''Stop
'        'Passa por cada uma das [ tabelas ] adicionadas ao dict [ dictTempDict ]
'        ' e confirma se o [ campo ] ora avaliado existe em pelo menos uma delas
'        Set db = CurrentDb
'        For Each vKey In dictTempDict
'
'            If gBbEnableErrorHandler Then On Error Resume Next
'            For Each vKeyField In db.TableDefs(vKey).Fields
''Stop
'            If Err.Number = 3265 Then Err = 0
'
'            'Debug.Print vKey
'                If vKeyField.Name = sQryField Then GetFldInQryGridTbls.bFoundQryFld = True: GoTo QryFldFound
''Stop
'            Next vKeyField
''Stop
'        Next vKey
''Stop
'
''MsgBox "tabledef concluído"
''Stop
'
'    End If
'    '--------------------------------------------------------------------------------------------
'
''Stop
'
'    'Se não tiver encontrado [ sQryField ] em uma das tabelas da Consulta procura também nas Queries
'    If Not GetFldInQryGridTbls.bFoundQryFld Then
'
'        dictTempDict.RemoveAll
'
'        '--------------------------------------------------------------------------------------------
'        sRegExPattern = "\[(qry)_[0-9]*\([0-9]*\)[a-zA-ZçÇ0-9(\)-_]*\]"
'        regEx.pattern = sRegExPattern
'
'        'vA = regEx.Test(sSQLtablesString )
'
'        Set mcRegExMatchColl = regEx.Execute(sSQLtablesString)
'        'Debug.Print sSQLtablesString
''Stop
'        'Monta dicionário com todas as tabelas usadas no trecho [ FROM ] da consulta
'        For Each vRegExKey In mcRegExMatchColl
'            vA = vRegExKey.Value
'            If Not dictTempDict.Exists(vRegExKey.Value) Then dictTempDict.Add vRegExKey.Value, vRegExKey
''Stop
'        Next vRegExKey
''Stop
'        'Passa por cada uma das [ tabelas ] adicionadas ao dict [ dictTempDict ]
'        ' e confirma se o [ campo ] ora avaliado existe em pelo menos uma delas
'        'Set db = CurrentDb
'        For Each vKey In dictTempDict
'
'
'            For Each vKeyField In CurrentDb.QueryDefs(vKey).Fields
'
'                vA = vKeyField.Name
'                If vA = sQryField Then GetFldInQryGridTbls.bFoundQryFld = True: GoTo QryFldFound
''Stop
'            Next vKeyField
''Stop
'
'        Next vKey
'        '--------------------------------------------------------------------------------------------
'
'    End If
'
''MsgBox "qrydef concluído"
''Stop
'
'QryFldFound:
'
'    dictTempDict.RemoveAll
'
'End Function
