Attribute VB_Name = "_Testes3 LockRec"
Option Compare Database
Option Explicit

Public Sub EditLockRecord(fForM As Form, lngCoD As Long, sTbE As String, sKeyField As String, bLockRcrd As Boolean)
    Dim vA, vB, vC
    'Dim iFrmIndexID As Integer
    Dim rTbE As DAO.Recordset
    Dim qDef As DAO.QueryDef
    
    Dim sTbELookUp As String
    
    Dim sStoredFile As String
    Dim sStoredFileFullPath As String
    Dim sMsgboxLine1 As String, sMsgboxLine2 As String
    Dim fSubFrm As Form
    Dim bBoL As Boolean
    Dim iWhere As Integer
    Dim sStartSqL As String
    Dim sNewSqL As String
    
    'qry_00(4)aRecEditLock
    'Bloqueia/Libera o registro pra edição
    
Stop
    'a Edição foi iniciada. Trava o registro para Edição por outro usuário
    If bLockRcrd Then
        Set rTbE = CurrentDb.OpenRecordset(sTbE, dbOpenDynaset, dbAppendOnly)
        rTbE.AddNew
        
        rTbE("SolicitanteIDfk") = lngCoD
        'rTbE.Fields("SolicitanteIDfk") = lngCoD
        
        'Retorna o código indexador do usuário
        vA = Val(sGbUserLogin)
        vA = "6220" 'sGbUserLogin
        sTbELookUp = "tbl_00(00)aSysUsers"
        vB = GetRecordCoD(vA, sTbELookUp, "UserLogin", "UserID")
'Stop
        rTbE.Fields("UserLoginIDfk") = vB
        'rTbE.Fields("StartTime") = Now()
        
        rTbE.Update
        rTbE.MoveLast
        
    'a Edição foi finalizada. Libera o registro para Edição por outro usuário
    Else
'Stop
        'atribui a Consulta/Tabela [ sTbE ] à variável qDef
        Set qDef = CurrentDb.QueryDefs(sTbE)
        
        'retorna o SQL da Consulta/Tabela
        sStartSqL = qDef.sql
        'Debug.Print sStartSqL
        
        'remove o [ ; ] ao final do SQL
        iWhere = InStrRev(sStartSqL, ";")
        sStartSqL = IIf(iWhere > 0, Mid(sStartSqL, 1, iWhere - 1), sStartSqL)
        'Debug.Print sStartSqL
        
        'atualiza o SQL com a filtragem apropriada pra buscar os registros desejados
        vA = "[" & sKeyField & "]" & " LIKE " & lngCoD
        vB = "WHERE " & vA
        sNewSqL = sStartSqL & vbNewLine & vB
        'Debug.Print sNewSqL
        Set qDef = Nothing
'Stop
        'atribui à variável rTbE o recordset da consulta atualizada para recuperação dos dados
        Set rTbE = CurrentDb.OpenRecordset(sNewSqL, dbOpenDynaset, dbSeeChanges)
        
        ''move pro último registro do Recordset pra retornar a contagem de registros
        'rTbE.MoveLast
        'vB = rTbE.RecordCount
        
'        rTbE.FindFirst vA
'        If rTbE.NoMatch Then
'Stop
'
'        End If
        'se houver registros
        If Not rTbE.BOF Then
            
            'garante que o apontador está no primeiro registro
            rTbE.MoveFirst
            
            'apaga todos os registros que atendam ao filtro
            Do Until rTbE.EOF
                rTbE.Delete
                rTbE.MoveNext
            Loop
        
        End If

        
    End If

'Stop
    rTbE.Close
    Set rTbE = Nothing
'Stop
    
    
    If (Err.Number <> 0) Then
        If (Err.Number = 2046) Then
        
        End If
    
    End If
    
    'usar esse trecho apenas pra testes durante o desenvolvimento
    ' não é necessário em tempo de execução
    
    'confirmar se a consulta sTbE está aberta no Banco de Dados
    bBoL = CurrentData.AllQueries(sTbE).IsLoaded
    If bBoL Then
        DoCmd.SelectObject acQuery, sTbE, True
        
        If gBbEnableErrorHandler Then On Error Resume Next
        DoCmd.Requery
    
    End If
    On Error GoTo -1

'Stop
    'Set rTbE = CurrentDb.OpenRecordset(sTbE, dbOpenDynaset, dbSeeChanges)

End Sub
