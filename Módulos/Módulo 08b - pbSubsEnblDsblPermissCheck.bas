Attribute VB_Name = "M�dulo 08b - pbSubsEnblDsblPermissCheck"
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
    
    'Abre consulta pra verificar se o [ Controle ] exige permiss�o pra ser habilitado e qual � a permiss�o exigida
    ' Se n�o exigir permiss�o libera o [ Controle ] para ser habilitado
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
    'Passa por todos os n�veis requeridos pelo Controle [ iRqrdPrmissLvlDscrpt ] existentes na
    ' tabela de permiss�es [ qry_00(3)aSysRqrdPrms ]
    ' e verifica se o [ controle ] ora analisado exige permiss�o espec�fica para ser acessado
    If Not (rsTbE.EOF And rsTbE.BOF) Then
        rsTbE.MoveFirst
        Do Until rsTbE.EOF = True
            iRqrdPrmissLvlDscrpt = Int(rsTbE.Fields("UserLoginLevlsIDfk"))

            'Recupera os n�veis de acesso do usu�rio [ iUserPrmissLvlDscrpt ] e compara cada um deles
            ' com o n�vel de acesso requerido [ iRqrdPrmissLvlDscrpt ] ora avaliado pra liberar ou n�o o acesso

            iKey = 1
            
            For Each vKey In clObjUserParams.dictUserPermissions
                iUserPrmissLvlDscrpt = vKey
                
                    
                'Verifica se as condi��es para libera��o do Controle foram atendidas
                ' .n�vel do usu�rio [ iUserPrmissLvlDscrpt ] menor que 10
                '   .. o usu�rio est� dentro da faixa de usu�rios com acesso elevado
                '
                ' .n�vel requerido [ iRqrdPrmissLvlDscrpt ] deve ser de grau inferior ao n�vel do usu�rio
                '   .. ou seja, o n�mero do n�vel deve ser maior ou igual ao n�mero do n�vel do usu�rio
                '
                'ex.: user  2      required 3 (d� acesso)
                '     user  3      required 2 (n�o d� acesso)
                If iUserPrmissLvlDscrpt < 10 Then
                    If iRqrdPrmissLvlDscrpt >= iUserPrmissLvlDscrpt Then bPrmissGranted = True
                                        
                Else
                    'Se o n�vel de permiss�o requerido [ iRqrdPrmissLvlDscrpt ] � igual ao n�vel do usu�rio [iUserPrmissLvlDscrpt]: libera o acesso.
                    If iUserPrmissLvlDscrpt = iRqrdPrmissLvlDscrpt Then bPrmissGranted = True
                    
                End If
                
                If Not bPrmissGranted Then
                                  
                 'Defini��o da dica do controle caso seja o acesso seja negado -------------------------------------------------------------------------------------
                 ' primeiro bloco exibe os dados do usu�rio logado: tipo da permiss�o e o nome do usu�rio
                   
                   
'Stop
                'Verifica se o usu�rio possue mais de uma permiss�o
                If clObjUserParams.dictUserPermissions.Count > 1 Then
                
                    
                    sCtrlTipTxtBuild1 = "As permiss�es do usu�rio [" & clObjUserParams.sUserName & " ]"
                    
                    sCtrlTipTxtBuild2 = sCtrlTipTxtBuild2 & Chr(13) & " [ " & vKey & " - " & clObjUserParams.dictUserPermissions(vKey) & " ]"
                    If gBbDebugOn Then Debug.Print sCtrlTipTxtBuild2
                          
                    sCtrlTipTxtBuild3 = Chr(13) & "n�o d�o acesso a essa funcionalidade" & _
                         Chr(13) & "                                - - " & Chr(13) & "N�vel de permiss�o requerido:"
                    
                    sCtrlTipTxtFrstSctn = sCtrlTipTxtBuild1 & sCtrlTipTxtBuild2 & sCtrlTipTxtBuild3
                    If gBbDebugOn Then Debug.Print " ---------------- "
                    If gBbDebugOn Then Debug.Print sCtrlTipTxtBuild1
                          
                Else
                    sCtrlTipTxtFrstSctn = "A permiss�o [ " & vKey & " - " & clObjUserParams.dictUserPermissions(vKey) & " ]  do usu�rio [ " & clObjUserParams.sUserName & " ]" & _
                         Chr(13) & "n�o d� acesso a essa funcionalidade" & _
                         Chr(13) & "                                - - " & Chr(13) & "N�vel de permiss�o requerido:"
                         
                End If
                
                   
                'Caso a permiss�o requerida seja inferior a de administrador,
                    ' exibe a rela��o de permiss�es que permitem o acesso ao Controle
                    If iRqrdPrmissLvlDscrpt > 0 Then
                        vB = Chr(13) & "  [ " & rsTbE.Fields("UserLoginLevlsIDfk") & " - " & rsTbE.Fields("UserLoginLevelDescri�") & " ]" & " ou"
                        
                        'Outros perfis que d�o acesso
                        If iRqrdPrmissLvlDscrpt < 10 Then
                            vC = Chr(13) & "  [ " & 0 & " a " & iRqrdPrmissLvlDscrpt & " ], dentro da faixa de acessos elevados"
                            
                        Else
                            vC = Chr(13) & "  [ " & 0 & " a " & "9" & " ], dentro da faixa de acessos elevados"
                            
                        End If
                        
                    Else
                    'Caso somente o Administrador possua permiss�o
                        vB = Chr(13) & "  [ 0 - Administrador ] "
                        vC = ""
                    End If
                    
                    sPrmissDeniedTipText = sCtrlTipTxtFrstSctn & vB & vC
                    '-------------------------------------------------------------------------------------------------------------------------------------------------
                End If
                'Se os crit�rios pra acesso ao Controle n�o tiverem sido atendidos n�o libera o acesso
                ' e exibe, na dica do Controle, o motivo.
                iKey = iKey + 1
            Next vKey
'Stop

            rsTbE.MoveNext
        
        Loop
    
        bCheckUserPermissionLevel.bPermissionGrated = bPrmissGranted
        bCheckUserPermissionLevel.sCtrlNewTipText = sPrmissDeniedTipText
    
    'Caso o [ Controle ] n�o tenha sido encontrado na tabela de permiss�es de acesso a controle [ qry_00(00)cSysRqrdPrmss ]
    ' significa que n�o h� restri��o para acess�-lo
    Else
        bCheckUserPermissionLevel.bPermissionGrated = True
            
    End If
    
    rsTbE.Close 'Close the recordset
    Set rsTbE = Nothing 'Clean up
'Stop
End Function
