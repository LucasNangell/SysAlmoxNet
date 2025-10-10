Attribute VB_Name = "_Testes_LstboxMultSelect b"
Option Compare Database
Option Explicit





Private Sub chkMontarGrupoORIGINAL_Click()
    Dim vA, vB
    Dim bMontarOpGrps As Boolean
    Dim cCtL As Control
    Dim cLstBoxCtrl As Control
    Dim sLstBoxCtrl As String
    Dim vKey As Variant
    Dim iWhere As Integer
    Dim sTagStr As String
    Dim iFrmIndexID As Integer
    Dim iItensSelecionados As Integer
    Dim iLstBoxFrstSelected As Integer
    Dim vListItem As Variant
    Dim iCounter As Integer
    Dim dDicT As Dictionary
    
'Stop
    'vA = ActiveControl.Name
    bMontarOpGrps = ActiveControl.Value
    iFrmIndexID = dictFormsSeqID(Me.Name)     'obtém o índice identificador do formulário que está sendo tratado
    
    'localiza dentro da TAG do controle checkbox pressionado o nome do controle listbox
    ' que deve sofrer a alteração multi select/single select
    sTagStr = Me.ActiveControl.Tag
    iWhere = InStr(1, sTagStr, "TgtCtl.")
    vB = Len(sTagStr) - Len("TgtCtl.")
    sLstBoxCtrl = Mid(sTagStr, iWhere + Len("TgtCtl."), vB)
'Stop  'Ctrl point
    
    If Len(sLstBoxCtrl) > 0 Then
        Set dDicT = dictListboxCtrls(iFrmIndexID)
        If dictListboxTgtCtrl(iFrmIndexID).Exists(sLstBoxCtrl) = True Then  'verifica se o controle já foi armazenado
            Set ListboxNav = dDicT(sLstBoxCtrl)
'        If dictListboxTgtCtrl(iFrmIndexID).Exists(sLstBoxCtrl) = True Then  'verifica se o controle já foi armazenado
'            Set ListboxSele'ct = dictListboxTgtCtrl(iFrmIndexID)(sLstBoxCtrl)
'Stop
        End If
        
        Set cLstBoxCtrl = Me.Controls(sLstBoxCtrl)
        vB = cLstBoxCtrl.Name
        cLstBoxCtrl.SetFocus
'Stop
    Else
       MsgBox "Não foi possível localizar o Listbox [ " & "teste" & " ] para alterar seu multiselect" & _
       vbCr & "Não será possível permitir multiselect nesse listbox", vbExclamation + vbOKOnly
        
    End If
'Stop
    'indica que o List Box de exibição de resultados deve ser tratado como Single Select
    ' e grava essa informação na classe respectiva
    ' esse parâmetro então é consultado a cada vez que ocorre o evento AfterUpdate da classe
    ' pra que o listbox se comporte como devido
    If bMontarOpGrps Then
        vA = ListboxNav.bListBoxMultSlct
        ListboxNav.bListBoxMultSlct = True

    Else
'Stop
        vA = ListboxNav.bListBoxMultSlct
        ListboxNav.bListBoxMultSlct = False
        cLstBoxCtrl.SetFocus

'MsgBox "stop", vbOKOnly
'Stop   '-------------------
        'garante que haja somente um item selecionado no controle alvo de exibição de resultados principal
        Set cCtL = cLstBoxCtrl
        iItensSelecionados = cCtL.ItemsSelected.Count
        
        iCounter = 0
        If iItensSelecionados > 1 Then
            For Each vListItem In cCtL.ItemsSelected   '-1 indica que o item está selecionado, 0 indica que não está
                iCounter = iCounter + 1
                'vA = vListItem
                
                If iCounter = 1 Then
                    cCtL.Selected(vListItem) = True
                    iLstBoxFrstSelected = vListItem
                Else
                    cCtL.Selected(vListItem) = False
                
                End If

            Next vListItem
            'cCtL.ListIndex = 0
'Stop
            'cCtL.Selected(0) = True
            cCtL.ListIndex = iLstBoxFrstSelected
'Stop
        End If
'Stop
    End If
'Stop
End Sub

Sub lstOPs_AfterUpdate()
    Dim vA, vB
    Dim iFrmIndexID As Integer
    Dim cCtL As Control
    Dim bBoL As Boolean
    Dim bCtrlFndInDict As Boolean
    Dim sMsgboxLine1 As String, sMsgboxLine2 As String
    Dim lgSlctdItems As Long
    Dim oClsObject As Object
    Dim sLckdBtnTipText As String
    Dim lgSlctdItem As Long, lgSlctdItemBse1 As Long, lgItemsCount As Long

    Set cCtL = ActiveControl
    iFrmIndexID = dictFormsSeqID(Me.Name)
    
    If dictColorCtls(iFrmIndexID).Exists(cCtL.Name) = True Then
        Set CurrControlColor = dictColorCtls(iFrmIndexID)(cCtL.Name)
        bCtrlFndInDict = True
    
    End If
'Stop
    'se o evento não tiver sido chamado manualmente pelo evento do controle na Classe
    ' informa ao [ dictColorCtls ] se o Controle deve ou não executar a rotina forçada de Update
    ' com alteração de outros controles do form a partir do AfterUpdate deste controle
    If Not bgbCallingByName Then
        'indica que todo o código da rotina AfterUpdate do controle deve ser executado
        If bCtrlFndInDict Then CurrControlColor.bForceAfterUpdate = True
        
    End If

'Stop
    '-------------------------------------
    '-------------------------------------------------------------------------------------------------
    'este trecho só deve ser executado se tiver sido chamado manualmente a partir
    ' do evento evento AfterUpdate do controle na Classe
    If bgbCallingByName Then
        
        lgSlctdItems = cCtL.ItemsSelected.Count
        bBoL = cCtL.Selected(cCtL.ListIndex)
'Stop
        If bBoL Then
            'Habilita os botões de operação com OP
            Call EnableCtrl(iFrmIndexID, Me!btnExibirOP, True)
            'If bCtrlFndInDict Then Me!btnExibirOP.ControlTipText = CurrControlColor.sOriginalTipText
'Stop
            Call EnableCtrl(iFrmIndexID, Me!btnExcluirOP, True)
            'If bCtrlFndInDict Then Me!btnExcluirOP.ControlTipText = CurrControlColor.sOriginalTipText
'Stop
            Call EnableCtrl(iFrmIndexID, Me!btnDuplicarOP, True)
            'If bCtrlFndInDict Then Me!btnDuplicarOP.ControlTipText = CurrControlColor.sOriginalTipText
'Stop
            Call EnableCtrl(iFrmIndexID, Me!btnPastaOP, True)
            'If bCtrlFndInDict Then Me!btnPastaOP.ControlTipText = CurrControlColor.sOriginalTipText
'Stop
            Call EnableCtrl(iFrmIndexID, Me!btnCalcularLombada, True)
            'If bCtrlFndInDict Then Me!btnCalcularLombada.ControlTipText = CurrControlColor.sOriginalTipText
'Stop

        Else
'Stop
            'Desabilita os botões de operação com OP
            sLckdBtnTipText = "deve haver uma OP ativa pra ser exibida"
            Call EnableCtrl(iFrmIndexID, Me!btnExibirOP, False, sLckdBtnTipText)
            'Me!btnExibirOP.ControlTipText = "deve haver uma OP ativa pra ser exibida"
'Stop
            sLckdBtnTipText = "apenas a última OP pode ser excluída"
            Call EnableCtrl(iFrmIndexID, Me!btnExcluirOP, False, sLckdBtnTipText)
            'Me!btnExcluirOP.ControlTipText = "apenas a última OP pode ser excluída"
'Stop
            sLckdBtnTipText = "deve haver uma OP ativa pra ser duplicada"
            Call EnableCtrl(iFrmIndexID, Me!btnDuplicarOP, False, sLckdBtnTipText)
            'Me!btnDuplicarOP.ControlTipText = "deve haver uma OP ativa pra ser duplicada"
'Stop
            sLckdBtnTipText = "deve haver uma OP ativa pra acessar sua pasta"
            Call EnableCtrl(iFrmIndexID, Me!btnPastaOP, False, sLckdBtnTipText)
            'Me!btnPastaOP.ControlTipText = "deve haver uma OP ativa pra acessar sua pasta"
'Stop
            sLckdBtnTipText = "deve haver uma OP ativa para o cálculo de lombada"
            Call EnableCtrl(iFrmIndexID, Me!btnCalcularLombada, False, sLckdBtnTipText)
            'Me!btnCalcularLombada.ControlTipText = "deve haver uma OP ativa para o cálculo de lombada"
'Stop
        End If
'Stop
        '-------------------------------------
        '------------------------------------------------------------------------------
        'Trata separadamente o botão mover pasta
        If lgSlctdItems = 2 Then
'Stop
            Call EnableCtrl(iFrmIndexID, Me!btnMoverArquivos, True)
            'If bCtrlFndInDict Then Me!btnMoverArquivos.ControlTipText = CurrControlColor.sOriginalTipText
'Stop
        Else
            sLckdBtnTipText = "deve haver duas OPs selecionadas pra mover arquivos"
            Call EnableCtrl(iFrmIndexID, Me!btnMoverArquivos, False, sLckdBtnTipText)
            'Me!btnMoverArquivos.ControlTipText = "deve haver duas OPs selecionadas pra mover arquivos"
            
        End If
        '------------------------------------------------------------------------------
        '-------------------------------------
        
'Stop
        '-------------------------------------
        '------------------------------------------------------------------------------
        'Trata separadamente o botão Excluir OP
        vA = cCtL.Name
        lgSlctdItem = cCtL.ListIndex
        lgSlctdItemBse1 = lgSlctdItem + 1
        lgItemsCount = cCtL.ListCount
        
        'identifica a seleção ativa da lista, ou seja, o item recentemente selecionado
        ' pois quando a Lista está em MultiSelect, o clique sobre um item pode ter sido
        ' pra selecioná-lo ou pra desmarcá-lo
        vA = ChkLstSlctItem(iFrmIndexID, cCtL)
'Stop
        
        'Executa a rotina se houver uma seleção ativa
        If lgSlctdItemBse1 = lgItemsCount And vA <> -1 Then
'Stop 'habilita botão
            Call EnableCtrl(iFrmIndexID, Me!btnExcluirOP, True)
        
        Else
'Stop 'desabilita botão
            sLckdBtnTipText = "apenas a última OP pode ser excluída"
            Call EnableCtrl(iFrmIndexID, Me!btnExcluirOP, False, sLckdBtnTipText)
        
        End If
        '------------------------------------------------------------------------------
        '-------------------------------------
            
    
    
    End If
    '-------------------------------------------------------------------------------------------------
    '-------------------------------------

End Sub
