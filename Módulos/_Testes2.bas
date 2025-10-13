Attribute VB_Name = "_Testes2"
Option Compare Database
Option Explicit


Function GetSortedSQL(sSQLList As String, sColToOrder As String, sOrderMode As String)
    Dim sFrstSQL As String

    If InStr(sSQLList, "ORDER BY") > 0 Then sFrstSQL = Split(sSQLList, "ORDER BY")(0) Else sFrstSQL = sSQLList
    GetSortedSQL = sFrstSQL & "ORDER BY " & sColToOrder & " " & sOrderMode
    
End Function

Public Sub PbSbBuildDictFieldsInQryTrgtCtrl(sForM As String, cTrgtCtrl As Control, sFilGrp As String)
    Dim vA, vB, vC
    Dim sRegExPattern As String
    Dim sQuerY As String
    Dim regEx As New RegExp
    Dim mcRegExMatchColl As MatchCollection
    Dim vRegExKey As Variant
    Dim vKeyField As Variant
    Dim vKey As Variant
    Dim rsTbE As Recordset
    Dim sSQLTrgtCtrl As String
    Dim qDf As QueryDef
    Dim strQryName As String
    Dim fField As Field
    Dim sTrgtCtrl As String
    
    On Error Resume Next
    
    sSQLTrgtCtrl = cTrgtCtrl.RowSource
    sTrgtCtrl = cTrgtCtrl.Name
        
    If Left(sSQLTrgtCtrl, 7) <> "SELECT" Then
        strQryName = sSQLTrgtCtrl
        Set qDf = CurrentDb.QueryDefs(strQryName)
        sSQLTrgtCtrl = qDf.sql
    Else
        Set qDf = CurrentDb.QueryDefs(sSQLTrgtCtrl)
    End If
    
    strQryName = qDf.Name
    For Each fField In qDf.Fields
        
        If Not clObjTargtCtrlParam.dictQryFields.Exists(fField.Name) Then
            clObjTargtCtrlParam.dictQryFields.Add fField.Name, "Grid"
        End If
    Next fField

    regEx.Global = True
    regEx.IgnoreCase = False

    sRegExPattern = "\[(tbl)_[0-9]*\([0-9]*\)[a-zA-ZÁ«0-9(\)-_]*\]"
    regEx.pattern = sRegExPattern
        
    Set mcRegExMatchColl = regEx.Execute(sSQLTrgtCtrl)

    For Each vRegExKey In mcRegExMatchColl
        If Not dictTempDict.Exists(vRegExKey.Value) Then dictTempDict.Add vRegExKey.Value, vRegExKey
    Next vRegExKey
    
    For Each vKey In dictTempDict
        For Each fField In CurrentDb.TableDefs(vKey).Fields
            If Err.Number = 3265 Then Err = 0
                If Not fField Is Nothing Then
                    If Not clObjTargtCtrlParam.dictQryFields.Exists(fField.Name) Then
                    clObjTargtCtrlParam.dictQryFields.Add fField.Name, "offGrid"
                End If
            End If
        Next fField

    Next vKey

    dictTempDict.RemoveAll
    
    sRegExPattern = "\[(qry)_[0-9]*\([0-9]*\)[a-zA-ZÁ«0-9(\)-_]*\]"
    regEx.pattern = sRegExPattern

    Set mcRegExMatchColl = regEx.Execute(sSQLTrgtCtrl)

    For Each vRegExKey In mcRegExMatchColl
        If Not dictTempDict.Exists(vRegExKey.Value) Then dictTempDict.Add vRegExKey.Value, vRegExKey
    Next vRegExKey

    For Each vKey In dictTempDict
        For Each fField In CurrentDb.QueryDefs(vKey).Fields
            If Not fField Is Nothing Then
                If Not clObjTargtCtrlParam.dictQryFields.Exists(fField.Name) Then
                    clObjTargtCtrlParam.dictQryFields.Add fField.Name, "offGrid"
                End If
            End If
        Next fField

    Next vKey
    
    dictTempDict.RemoveAll
    
End Sub


