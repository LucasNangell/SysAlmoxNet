Attribute VB_Name = "_Temp"
Option Compare Database
Option Explicit


Public Sub MskdTxtbox_TextMask(cTrgtTxtbox As Control, sCustomFormat As String, Optional sPrefix As String, Optional iFormatMaxLen As Integer)
    Dim vA, vB, vC
    Dim sCtrl As String
    Dim sForM As String
    Dim sActualText As String, sOldTxt As String
    Dim iTypedCharPos As Integer
    Dim sLastTypedChar As String
    Dim bUseDecimal As Boolean
    Dim iSelStartIncrement As Integer
    Dim bChangeDecimal As Boolean
    Dim iQtCasaDecimal As Integer
    Dim sInteiro As String, sDecimal As String
    Dim iInT As Integer
    Dim bMskdCtrl As Boolean
    Dim sCleanTxt As String
    Dim sTxt As String
    Dim sDt1 As String, sDt2 As String, sDt3 As String

'------------- Chamada da Função  ---------------------------------------------
'     Call MskdTxtbox_TextMask(ActiveControl, "P_###,###", , 6)
'     Call MskdTxtbox_TextMask(ActiveControl, "#,###.00", "R$ ", 6)
'------------------------------------------------------------------------------

    sCtrl = cTrgtTxtbox.Name
    sForM = cTrgtTxtbox.Parent.Name

    'Confirma se o controle tem [ bMskdCtrl = TRUE ]
    ' verifica se ele existe no dict [ dictCtrlBehvrParams(sForM) ]
    If IsObject(dictCtrlBehvrParams(sForM)) Then
        If dictCtrlBehvrParams(sForM).Exists(sCtrl) = True Then
            Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sCtrl)
            bMskdCtrl = clObjCtrlBehvrParams.bMskdCtrl
        
        End If
    
    End If
    
    If Not bMskdCtrl Then Exit Sub

    On Error Resume Next
    sActualText = cTrgtTxtbox.Text
    If Err.Number = 2185 Then sActualText = cTrgtTxtbox.Value
    On Error GoTo -1
    iTypedCharPos = cTrgtTxtbox.SelStart
    
    'Se o controle estiver VAZIO sai da rotina
    If sActualText = "" Then Exit Sub
    
    'Recupera o texto anterior do controle
    If Not IsNull(cTrgtTxtbox.OldValue) Then sOldTxt = cTrgtTxtbox.OldValue

    sLastTypedChar = Mid(sActualText, iTypedCharPos, 1)
    
    'Se houver [.] na [ sCustomFormat ] passada, atribui TRUE para [ bUseDecimal ]
    bUseDecimal = InStr(sCustomFormat, ".") > 0
    
    If InStr(sActualText, sPrefix) = 0 Then iSelStartIncrement = Len(sPrefix)
    
    If bUseDecimal Then
        'Recupera a quantidade de casas decimais do formato solicitado
        iQtCasaDecimal = Len(sCustomFormat) - InStr(sCustomFormat, ".")
        
        'Se o texto atual tiver [ Vírgula ] verifica se as alterações deverão ser feitas
        ' no número inteiro ou nas casas decimais
        If Len(sActualText) > 1 Then
            
            If InStr(sActualText, ",") > 0 Then
                vB = Len(sActualText) - iTypedCharPos
                'Se o cursor do mouse estiver após a [ Vírgula ] indica que as alterações
                ' serão nas casas decimais
                If vB <= iQtCasaDecimal Then bChangeDecimal = True
            
            'Se o texto atual não tiver [ Vírgula ] mas o texto anterior sim, indica que o usuário está apagando
            ' então o valor original é mantido e o cursor é posicionado antes da vírgula
            ElseIf InStr(sActualText, ",") = 0 And InStr(sOldTxt, ",") > 0 Then
                sActualText = sOldTxt
                'iSelStartIncrement = -1
            End If
            
        End If
        'Caso o usuário tenha digitado [ Vírgula ], a segunda é removida
        sActualText = Replace(sActualText, ",,", ",")
        
        'Percorre cada caracter de [ sActualText ] para montar [ sInteiro ] e [ sDecimal ]
        For iInT = 1 To Len(sActualText)
            
            'A cada caracter analisado verifica-se se é um número ou se é uma [ Vírgula ]
            ' sempre que for um número esse valor é adicionado a [ sInteiro ] até o momento
            ' em que o código passar pela [ Vírgula ], a partir daí [ sDecimal ] começa a ser montada
            If Mid(sActualText, iInT, 1) = "," Then vA = True
            If IsNumeric(Mid(sActualText, iInT, 1)) Then
                If vA = False Then
                    sInteiro = sInteiro & Mid(sActualText, iInT, 1)
                Else
                    sDecimal = sDecimal & Mid(sActualText, iInT, 1)
                End If
            End If
        Next iInT
        
        'Verifica se a quantidade de caracteres de [ sDecimal ] é compatível com a quantidade estabelecida por [ sCustomFormat ]
        ' se a quantidade é maior, apenas os caracteres da esquerda são aproveitados
        If Len(sDecimal) > iQtCasaDecimal Then
            sDecimal = Left(sDecimal, iQtCasaDecimal)
        Else
            'Completa a quantidade de casas decimais com ZERO a direita caso necessário
            Do While Len(sDecimal) < iQtCasaDecimal: sDecimal = sDecimal & "0": Loop
        End If
        
        'Atribui a [ sActualText ] o texto formatado com o prefixo.
        sActualText = sPrefix & Format(sInteiro & "," & sDecimal, sCustomFormat)
        
        If Len(sActualText) > Len(sOldTxt) + 1 Then
            iSelStartIncrement = 1
        ElseIf Len(sActualText) < Len(sOldTxt) - 1 Then
            iSelStartIncrement = -1
        End If
            
        cTrgtTxtbox.Value = sActualText
        
        cTrgtTxtbox.SelStart = iTypedCharPos + iSelStartIncrement

    
    Else 'Caso não se trate de uma máscara com número decimal
        
        'Limpa o controle pra deixar apenas números
        For iInT = 1 To Len(sActualText)
            sTxt = Mid(sActualText, iInT, 1)
            If IsNumeric(sTxt) Then sCleanTxt = sCleanTxt & sTxt
        
        Next iInT
        
        If Len(sCleanTxt) > iFormatMaxLen Then
            'Volta ao valor anterior
            cTrgtTxtbox.Undo
            Exit Sub
        End If
        
        'Se for um campo de data adapta [ sCustomFormat ] de acordo com a quantidade de caracters digitados
        If sCustomFormat = "##/##/####" Then
            Select Case Len(sCleanTxt)
                Case 1: sCustomFormat = IIf(Int(sCleanTxt) > 3, "0#/", "0")
                Case 2: sCustomFormat = IIf(Right(sActualText, 1) = "/", "0#/", "0#")
                Case 3: sCustomFormat = IIf(Int(Right(sCleanTxt, 1)) > 1, "0#/0#/", "0#/#")
                Case 4: sCustomFormat = IIf(Right(sActualText, 1) = "/", "0#/0#/", "0#/0#")
                Case 5: sCustomFormat = "00/00/0"
                Case 6: sCustomFormat = IIf(Int(Right(sCleanTxt, 2)) <> 20, "0#/0#/####", "0#/0#/##"): _
                If Int(Right(sCleanTxt, 2)) <> 20 Then sCleanTxt = Left(sCleanTxt, 4) & "20" & Right(sCleanTxt, 2)
                Case 7: sCustomFormat = "0#/0#/###"
                Case 8: sCustomFormat = "0#/0#/####"
            End Select
        End If

        sActualText = sCleanTxt
        sActualText = sPrefix & Format(sActualText, sCustomFormat)
        
        cTrgtTxtbox.Value = sActualText
        cTrgtTxtbox.SelStart = cTrgtTxtbox.SelLength
        
    End If
    
End Sub

