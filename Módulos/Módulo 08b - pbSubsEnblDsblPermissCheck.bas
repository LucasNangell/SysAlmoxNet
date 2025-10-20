Attribute VB_Name = "Módulo 08b - pbSubsEnblDsblPermissCheck"
Option Compare Database
Option Explicit


Public Function bCheckUserPermissionLevel(sForM As String, sTweakableCtrL As String) As vCtrlPrmissGrnted
    Dim vA, vB, vC, vD
    Dim rsTbE As Recordset
    Dim sQuerY As String
    Dim sWhere As String
    Dim lngFoundRecs As Long
    Dim vKey As Variant
    Dim iKey As Integer
    Dim iUserPrmissLvlDscrpt As Integer
    Dim iRqrdPrmissLvlDscrpt As Integer
    Dim sPrmissDeniedTipText As String
    Dim bPrmissGranted As Boolean
    Dim sCtrlTipTxtBuild1 As String, sCtrlTipTxtBuild2 As String, sCtrlTipTxtBuild3 As String, sCtrlTipTxtFrstSctn As String
    
    'Abre consulta pra verificar se o [ Controle ] exige permissão pra ser habilitado e qual é a permissão exigida
    ' Se não exigir permissão libera o [ Controle ] para ser habilitado
    sQuerY = "qry_00(00)cSysRqrdPrmss"
    sWhere = "([ForM] Like " & """" & sForM & """" & ") And ([Control] Like " & """" & sTweakableCtrL & """" & ")"
    If gBbDebugOn Then Debug.Print sWhere

    Set rsTbE = CurrentDb.OpenRecordset(sQuerY, dbOpenDynaset, dbReadOnly)
    rsTbE.Filter = sWhere
    Set rsTbE = rsTbE.OpenRecordset

    lngFoundRecs = rsTbE.RecordCount
    If lngFoundRecs > 0 Then
        rsTbE.MoveLast
        lngFoundRecs = rsTbE.RecordCount
    
    End If

'If gBbDepurandoLv01b Then MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Checar rsTbE"
'Stop
    'Passa por todos os níveis requeridos pelo Controle [ iRqrdPrmissLvlDscrpt ] existentes na
    ' tabela de permissões [ qry_00(3)aSysRqrdPrms ]
    ' e verifica se o [ controle ] ora analisado exige permissão específica para ser acessado
    If Not (rsTbE.EOF And rsTbE.BOF) Then
        rsTbE.MoveFirst
        Do Until rsTbE.EOF = True
            iRqrdPrmissLvlDscrpt = Int(rsTbE.Fields("UserLoginLevlsIDfk"))

            'Recupera os níveis de acesso do usuário [ iUserPrmissLvlDscrpt ] e compara cada um deles
            ' com o nível de acesso requerido [ iRqrdPrmissLvlDscrpt ] ora avaliado pra liberar ou não o acesso

            iKey = 1
            
            For Each vKey In clObjUserParams.dictUserPermissions
                iUserPrmissLvlDscrpt = vKey
                
                    
                'Verifica se as condições para liberação do Controle foram atendidas
                ' .nível do usuário [ iUserPrmissLvlDscrpt ] menor que 10
                '   .. o usuário está dentro da faixa de usuários com acesso elevado
                '
                ' .nível requerido [ iRqrdPrmissLvlDscrpt ] deve ser de grau inferior ao nível do usuário
                '   .. ou seja, o número do nível deve ser maior ou igual ao número do nível do usuário
                '
                'ex.: user  2      required 3 (dá acesso)
                '     user  3      required 2 (não dá acesso)
                If iUserPrmissLvlDscrpt < 10 Then
                    If iRqrdPrmissLvlDscrpt >= iUserPrmissLvlDscrpt Then bPrmissGranted = True
                                        
                Else
                    'Se o nível de permissão requerido [ iRqrdPrmissLvlDscrpt ] é igual ao nível do usuário [iUserPrmissLvlDscrpt]: libera o acesso.
                    If iUserPrmissLvlDscrpt = iRqrdPrmissLvlDscrpt Then bPrmissGranted = True
                    
                End If
                
                If Not bPrmissGranted Then
                                  
                 'Definição da dica do controle caso seja o acesso seja negado -------------------------------------------------------------------------------------
                 ' primeiro bloco exibe os dados do usuário logado: tipo da permissão e o nome do usuário
                   
                   
'Stop
                'Verifica se o usuário possue mais de uma permissão
                If clObjUserParams.dictUserPermissions.Count > 1 Then
                
                    
                    sCtrlTipTxtBuild1 = "As permissões do usuário [" & clObjUserParams.sUserName & " ]"
                    
                    sCtrlTipTxtBuild2 = sCtrlTipTxtBuild2 & Chr(13) & " [ " & vKey & " - " & clObjUserParams.dictUserPermissions(vKey) & " ]"
                    If gBbDebugOn Then Debug.Print sCtrlTipTxtBuild2
                          
                    sCtrlTipTxtBuild3 = Chr(13) & "não dão acesso a essa funcionalidade" & _
                         Chr(13) & "                                - - " & Chr(13) & "Nível de permissão requerido:"
                    
                    sCtrlTipTxtFrstSctn = sCtrlTipTxtBuild1 & sCtrlTipTxtBuild2 & sCtrlTipTxtBuild3
                    If gBbDebugOn Then Debug.Print " ---------------- "
                    If gBbDebugOn Then Debug.Print sCtrlTipTxtBuild1
                          
                Else
                    sCtrlTipTxtFrstSctn = "A permissão [ " & vKey & " - " & clObjUserParams.dictUserPermissions(vKey) & " ]  do usuário [ " & clObjUserParams.sUserName & " ]" & _
                         Chr(13) & "não dá acesso a essa funcionalidade" & _
                         Chr(13) & "                                - - " & Chr(13) & "Nível de permissão requerido:"
                         
                End If
                
                   
                'Caso a permissão requerida seja inferior a de administrador,
                    ' exibe a relação de permissões que permitem o acesso ao Controle
                    If iRqrdPrmissLvlDscrpt > 0 Then
                        vB = Chr(13) & "  [ " & rsTbE.Fields("UserLoginLevlsIDfk") & " - " & rsTbE.Fields("UserLoginLevelDescriç") & " ]" & " ou"
                        
                        'Outros perfis que dão acesso
                        If iRqrdPrmissLvlDscrpt < 10 Then
                            vC = Chr(13) & "  [ " & 0 & " a " & iRqrdPrmissLvlDscrpt & " ], dentro da faixa de acessos elevados"
                            
                        Else
                            vC = Chr(13) & "  [ " & 0 & " a " & "9" & " ], dentro da faixa de acessos elevados"
                            
                        End If
                        
                    Else
                    'Caso somente o Administrador possua permissão
                        vB = Chr(13) & "  [ 0 - Administrador ] "
                        vC = ""
                    End If
                    
                    sPrmissDeniedTipText = sCtrlTipTxtFrstSctn & vB & vC
                    '-------------------------------------------------------------------------------------------------------------------------------------------------
                End If
                'Se os critérios pra acesso ao Controle não tiverem sido atendidos não libera o acesso
                ' e exibe, na dica do Controle, o motivo.
                iKey = iKey + 1
            Next vKey
'Stop

            rsTbE.MoveNext
        
        Loop
    
        bCheckUserPermissionLevel.bPermissionGrated = bPrmissGranted
        bCheckUserPermissionLevel.sCtrlNewTipText = sPrmissDeniedTipText
    
    'Caso o [ Controle ] não tenha sido encontrado na tabela de permissões de acesso a controle [ qry_00(00)cSysRqrdPrmss ]
    ' significa que não há restrição para acessá-lo
    Else
        bCheckUserPermissionLevel.bPermissionGrated = True
            
    End If
    
    rsTbE.Close 'Close the recordset
    Set rsTbE = Nothing 'Clean up
'Stop
End Function
