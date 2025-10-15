VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_00(1)bSysTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim clObjTimerFrmParams As New cls_09cParamsToOpenForms
Private dictFormsPosition As New Dictionary


Private Sub Form_Timer()
    
    If gBbEnableErrorHandler Then On Error Resume Next
    
'parei aqui1: confirmar a necessidade desse Resume Next e se o tratamento de erro no final deve permanecer. estava comentado

    Dim fForM As Form
    Dim clFormCords As cls_09zOpndFormsPositions
    Dim iInT As Integer
    Dim sTrgtForm As String
    Dim sTrggForM As String
    Dim lngDeslocHoriz As Long
    Dim lngDeslocVert As Long
    Dim bIsTrggForm As Boolean
    Dim iKey As Integer
    Dim iCountForm As Integer
    'Formulário carregado e oculto durante o uso do sistema pra
    ' monitoramento da posição dos formulários abertos de forma a possibilitar
    ' que os formulários se movam juntos quando ancorados
    
    'Se a quantidade de formulários abertos é menor que 3
    ' remove os itens do [ dictFormsPosition ]
    If Forms.Count < 3 Then
        dictFormsPosition.RemoveAll
        dictFormsParams.RemoveAll
        DoCmd.Close
        Exit Sub
    End If

    
    With clObjTimerFrmParams
    
        For Each fForM In Forms
            .bIsTrggForm = False
            'Verifica se existe algum item no [ dictFormsParams ] de [ TrggForm ] não aberto
            ' Caso haja, remove o item
            For iInT = 0 To dictFormsParams.Count - 1
                Set clObjFormsParams = dictFormsParams(dictFormsParams.Keys(iInT))
                If clObjFormsParams.sTrggFormName = fForM.Name Then .bIsTrggForm = True
            Next iInT
            If Not .bIsTrggForm Then
                If dictFormsParams.Exists(fForM.Name) = True Then dictFormsParams.Remove (fForM.Name)
            End If
            'Verifica se existe algum item no [ dictFormsPosition ] de [ TrggForm ] não aberto
            ' Caso haja, remove o item
            For iInT = 0 To dictFormsPosition.Count - 1
                Set clFormCords = dictFormsPosition(dictFormsPosition.Keys(iInT))
                If Not CurrentProject.AllForms(clFormCords.sFormName).IsLoaded Then dictFormsPosition.Remove (fForM.Name)
            Next iInT
        Next fForM
        
        iCountForm = Forms.Count
        
        
        Do
        iCountForm = iCountForm - 1
        Set fForM = Forms(iCountForm)
        'For Each fForM In Forms
            'Abre a [ clFormCords ] de cada [ fForm ] aberto para verificar mudanças de posição
            If dictFormsPosition.Exists(fForM.Name) = True Then
                Set clFormCords = dictFormsPosition(fForM.Name)
                
                'Se houver mudança na posição do [ fForm ]
                If clFormCords.lngLeft <> fForM.WindowLeft And clFormCords.sFormName = fForM.Name Or _
                   clFormCords.lngTop <> fForM.WindowTop And clFormCords.sFormName = fForM.Name Then
Back:
                   
                   'Atualiza a classe com a nova posição
                   clFormCords.lngLeft = fForM.WindowLeft
                   clFormCords.lngTop = fForM.WindowTop
                   
                    'Percorre os itens do [ dictFormsParams ]
                    For iInT = 0 To dictFormsParams.Count - 1
                        Set clObjFormsParams = dictFormsParams(dictFormsParams.Keys(iInT))
                        'Caso o item seja referente a um formulário não aberto o item é removido do dicionário
                        If Not CurrentProject.AllForms(clObjFormsParams.sTrggFormName).IsLoaded Then
                            dictFormsParams.Remove (dictFormsParams.Keys(iInT))
                            iInT = iInT - 1
                        End If
                        
                        
                        If fForM.Name = clObjFormsParams.sTrggFormName Or fForM.Name = clObjFormsParams.sTrgtFormName Then
                            .sTrggForM = clObjFormsParams.sTrggFormName
                            .sTrgtForm = clObjFormsParams.sTrgtFormName
                            
                            'Caso o botão de ancoragem esteja pressionado
                            If clObjFormsParams.bTrgtFormIsDocked = True Then
                                
                                'Verifica se [ fForm ] é um [ TrggForm ] ou um [ TrgtForm ]
                                '  e atualiza os dados da classe [ clObjFormsParams ]
                                
                                If fForM.Name = .sTrggForM Then
                                    clObjFormsParams.lngTrggLeft = fForM.WindowLeft
                                    clObjFormsParams.lngTrggTop = fForM.WindowTop
                                    
                                    clObjFormsParams.lngTrgtTop = fForM.WindowTop + clObjFormsParams.lngTrgtVertOffset
                                    clObjFormsParams.lngTrgtLeft = fForM.WindowLeft + clObjFormsParams.lngTrgtHorizOffset
                                    .bIsTrggForm = False
                                    
                                    On Error GoTo -1
                                    
                                    'Executa a mudança de posição do [ TrgtForm ] vinculado
                                    Set clObjTimerFrmParams.fTrgtForm = Forms(clObjTimerFrmParams.sTrgtForm)
                                    Call FormLoad08_SetFormPosition(clObjTimerFrmParams)
                                
                                End If
                                If fForM.Name = .sTrgtForm Then
                                    clObjFormsParams.lngTrgtLeft = fForM.WindowLeft
                                    clObjFormsParams.lngTrgtTop = fForM.WindowTop
                                    
                                    clObjFormsParams.lngTrggTop = fForM.WindowTop - clObjFormsParams.lngTrgtVertOffset
                                    clObjFormsParams.lngTrggLeft = fForM.WindowLeft - clObjFormsParams.lngTrgtHorizOffset
                                    .bIsTrggForm = True
                                    
                                    On Error GoTo -1
                                    
                                    'Executa a mudança de posição do [ TrggForm ] vinculado
                                    Set clObjTimerFrmParams.fTrgtForm = Forms(clObjTimerFrmParams.sTrggForM)
                                    Call FormLoad08_SetFormPosition(clObjTimerFrmParams)
                                    
                                End If
    
                            End If
     
                        End If
    
                    Next iInT
                    
                End If
                
            Else
                'Caso o dicionário não tenha nenhum item para o [ fForm ]
                Set clFormCords = New cls_09zOpndFormsPositions
                dictFormsPosition.Add fForM.Name, clFormCords
                
                clFormCords.lngLeft = fForM.WindowLeft
                clFormCords.lngTop = fForM.WindowTop
                clFormCords.sFormName = fForM.Name
            End If
        
        'Next fForM
        Loop While iCountForm > 0
        'Percorre novamente todos os formulários abertos para verificar se é necessária mais uma mudança de posição
        For Each fForM In Forms
            
            If dictFormsPosition.Exists(fForM.Name) = True Then
                Set clFormCords = dictFormsPosition(fForM.Name)
                If clFormCords.sFormName <> fForM.Name Then
                    dictFormsPosition.Remove (fForM.Name)
                    
                Else
                    If clFormCords.lngLeft <> fForM.WindowLeft And clFormCords.sFormName = fForM.Name Or _
                       clFormCords.lngTop <> fForM.WindowTop And clFormCords.sFormName = fForM.Name Then
                       'Caso haja mudança de posição retorna a posição [ Back ]
                       GoTo Back
                    
                    End If
                End If
            End If
                    
    
        Next fForM
        
        'Atualiza a posição de todos os formulários abertos no [ dictFormsPosition ]
        For iKey = 0 To dictFormsPosition.Count - 1
            Set clFormCords = dictFormsPosition(dictFormsPosition.Keys(iKey))
            If CurrentProject.AllForms(clFormCords.sFormName).IsLoaded Then
                clFormCords.lngLeft = Forms(clFormCords.sFormName).WindowLeft
                clFormCords.lngTop = Forms(clFormCords.sFormName).WindowTop
            Else
                dictFormsPosition.Remove (dictFormsPosition.Keys(iKey))
            End If
    
        Next iKey
    
    End With
    
    If CurrentProject.AllForms("frm_00(1)bSysTimer").IsLoaded Then Me.TimerInterval = 10

End Sub

