Attribute VB_Name = "Módulo 00d - Aux (FormNav)"
Option Compare Database
Option Explicit

Public Sub HighlightClrChange(iCtrlType As Integer, cCtrL As Control, bLostFocus As Boolean)
    
    Dim vA, vB
    Dim sLockdStatus As String
    Dim sStR1 As String, sStR2 As String
    Dim sCtrL As String
    Dim sForM As String
    Dim bIsDirty As Boolean
    
    Dim bColorHighlight As Boolean
    Dim bOnDirty As Boolean
    Dim sCtlValue As Variant  'pode se referir a valores string ou a valores Integer, dependendo do controle analisado
    
    
    
    Dim iCtlItemsCount As Integer
    Dim lgInT As Integer
    Dim vItem As Variant
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Executa [ HighlightClrChange ]"
'Stop

    sCtrL = cCtrL.Name
    sForM = cCtrL.Parent.Name
    sLockdStatus = IIf(cCtrL.Locked Or Not cCtrL.Enabled, "Dsbld", "Enbld")
    Set clObjLckdStatusParam = Nothing
'Stop
    'Só faz a mudança de [ HighlightColor ] se o controle estiver [ habilitado ]
    If sLockdStatus = "Enbld" Then
        
        'Recupera os valores de cor do controle
        '---------------------------------------------------------------
        If IsObject(dictCtrlEnblDsblParams(sForM)) Then
        'If dictCtrlEnblDsblParams.Exists(sForM) = True Then
            If dictCtrlEnblDsblParams(sForM).Exists(sCtrL) = True Then
                Set clObjCtrlsEnblDsblParams = dictCtrlEnblDsblParams(sForM)(sCtrL)
                    If clObjCtrlsEnblDsblParams.dictParamByLckdStatus.Exists(sLockdStatus) = True Then
                        Set clObjLckdStatusParam = clObjCtrlsEnblDsblParams.dictParamByLckdStatus(sLockdStatus)
                
                    End If
            End If
        End If
        '---------------------------------------------------------------
'Stop
'Stop
        
        'Recupera os parâmetros [ BehvrParams ] do controle pra saber como ele deve mudar de cor
        ' à medida em que o usuário navega pelo form
        '---------------------------------------------------------------
        Set clObjCtrlBehvrParams = Nothing
        
        If IsObject(dictCtrlBehvrParams(sForM)) Then
        
            'Verifica se o Controle existe no [ dictCtrlBehvrParams(sForM) ]
            If dictCtrlBehvrParams(sForM).Exists(sCtrL) = True Then
                Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sCtrL)
                    
                bColorHighlight = clObjCtrlBehvrParams.bColorHighlight
                bOnDirty = clObjCtrlBehvrParams.bOnDirty
            
            Else
                Exit Sub
            
            End If
            
        End If
    
    
        Select Case iCtrlType
            Case acListBox
'Stop
                If bLostFocus Then
                    iCtlItemsCount = cCtrL.ItemsSelected.Count
    
                    lgInT = IIf((iCtlItemsCount < 1) Or (bOnDirty = False), 1, 2)
                    cCtrL.BackColor = IIf(lgInT = 1, clObjLckdStatusParam.lngLckdStatusBackColor, GbLngDIRTclrBackColor)
'Stop
                Else
                    cCtrL.BackColor = HexToLongRGB(GbLngHLclrBackColor)
    
                End If
            
            Case Else
'Stop
                'Se o controle tiver perdido o foco muda a cor para DIRTclr ou cor original do controle
                ' a depender do parâmetro [ bOnDirty ]
                If bLostFocus Then
'Stop
                    vA = cCtrL.Name
                    vA = bOnDirty
                    sCtlValue = cCtrL.Value
                    
                    bIsDirty = IIf((IsNull(sCtlValue) Or sCtlValue = "") Or (bOnDirty = False), False, True)
                    cCtrL.BackColor = IIf(bIsDirty, GbLngDIRTclrBackColor, clObjLckdStatusParam.lngLckdStatusBackColor)
                
                
                'Se o controle tiver recebido o foco muda a cor para HIGHLIGHTclr
                ' a depender do parâmetro [ bColorHighlight ]
                Else
                    
                    If bColorHighlight Then cCtrL.BackColor = GbLngHLclrBackColor
'Stop
                End If
        
        End Select
     
    End If

End Sub



Public Sub MskdTxtbox02_TextMask(cTxtboxCtrl As Control, sTypedText As String, sFormat As String, iFormatLen As Integer)

    Dim vA, vB
    Dim sTxtOldValue As String
    Dim lTxtValue As Long
    Dim sTxt As String
    Dim iInT As Integer
    Dim sTxtboxNrValue As String
    Dim sTxtboxNrTypedTxt As String
    Dim sCtrL As String
    Dim sForM As String
    Dim sStR1 As String, sStR2 As String
    Dim bMskdCtrl As Boolean
    
    'Dim iFormatLen As Integer
    'Dim sFormat As String, iFormatLen As Integer
    'Dim sTypedText As String, sTxtOldValue As String
   
    
    '-------------------------------------------------------------------------------------------------
    'Adiciona ao valor do controle, digitado pelo usuário, a máscara definida em [ sFormat ]
    '-------------------------------------------------------------------------------------------------
     
    'Chamada da função
    'Call MskdTxtbox02_TextMask(ActiveControl, ActiveControl.Text, "P_###,###", 6)
    
    sCtrL = cTxtboxCtrl.Name
    sForM = cTxtboxCtrl.Parent.Name
    
    vA = cTxtboxCtrl.Value
    vB = cTxtboxCtrl.Text
'Stop
    
    'Confirma se o controle tem [ bMskdCtrl = TRUE ]
    ' verifica se ele existe no dict [ dictCtrlBehvrParams(sForM) ]
    If IsObject(dictCtrlBehvrParams(sForM)) Then
    'dictCtrlBehvrParams(sForM)(sCtrL)
        If dictCtrlBehvrParams(sForM).Exists(sCtrL) = True Then
            Set clObjCtrlBehvrParams = dictCtrlBehvrParams(sForM)(sCtrL)
'Stop
            bMskdCtrl = clObjCtrlBehvrParams.bMskdCtrl
        
        End If
    
    End If
    
    If Not bMskdCtrl Then Exit Sub
'Stop
    
    'sFormat = "#,###,###"  ' "0,000,000"
    'iFormatLen = Len(Replace(sFormat, ",", ""))
    
   'vA = "12.345.67d"
    
    'sTypedText = Me!txtProcesso.Text
    If sTypedText = "" Then
        cTxtboxCtrl = ""
        'TxtboxWithMask = ""
        Exit Sub
        
    End If
    
    
    'Armazena o valor antes da digitação mais recente, que pode ter que ser descartada se
    ' tiver sido uma letra ou ultrapassado a quantidade de caracteres da máscara
    'vA = Val(CLng(cTxtboxCtrl))
    'vB = Str(vA)
    
'Stop
    If cTxtboxCtrl <> "" Then
        sTxt = cTxtboxCtrl
        'limpa o controle pra deixar apenas números
        sTxtOldValue = MskdTxtbox01_ClearNr(sTxt)
    
    End If
'Stop

    If sTypedText <> "" Then
        'Limpa o controle pra deixar apenas números
        sTxtboxNrTypedTxt = MskdTxtbox01_ClearNr(sTypedText)
        
    End If
'Stop
    
    
    '--------------------------------------------
    'Identifica se foram digitados apenas números
    ' se tiverem sido digitadas LETRAS dá erro
    On Error Resume Next
    If sTxtboxNrTypedTxt <> "" Then lTxtValue = Val(CLng(sTxtboxNrTypedTxt))
    
    If (Err.Number = 13) Then
        sTxt = sTxtOldValue
        'cTxtboxCtrl = sTxtOldValue
        'TxtboxWithMask = sTxtOldValue
    
    Else
        
        '----------------------------------------------------------
        'Identifica se a quantidade de dígitos ultrapassou o limite
'Stop
        'vA = Str(lTxtValue)
        'vB = Len(vA)
        
        vA = Val(lTxtValue)
        vB = Len(vA)
        If vB > iFormatLen Then
            sTxt = sTxtOldValue
            'cTxtboxCtrl = Format(sTxtOldValue, sFormat)
            'TxtboxWithMask = sTxtOldValue
            
        Else
            sTxt = lTxtValue
            'cTxtboxCtrl = Format(lTxtValue, sFormat)
            'TxtboxWithMask = Format(lTxtValue, sFormat)
            
        End If
        
        '----------------------------------------------------------
        
    End If
    On Error GoTo -1
    '--------------------------------------------
    
    cTxtboxCtrl = Format(sTxt, sFormat)
    cTxtboxCtrl.SelStart = cTxtboxCtrl.SelLength

End Sub

