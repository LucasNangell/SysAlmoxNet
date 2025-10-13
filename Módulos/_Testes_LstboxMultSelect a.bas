Attribute VB_Name = "_Testes_LstboxMultSelect a"
Option Compare Database
Option Explicit

Private Sub MultiSelect02_ToggleListbox()
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


Sub MultiSelect01_Start(cCtrL As Control)
    Dim vA, vB, vC
    Dim iWhere As Integer
    Dim sTagParams As String
    Dim sParam As String
    Dim vTagSectionParams As Variant
    Dim sTagSection As String
    Dim sFilGrp As String
    Dim sForM As String
Stop
    
'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "MultiSelectStart"
Stop
    
    sForM = cCtrL.Parent.Name
    
    'Confirma se o check é um controle tipo [ Target Multiselect ]
    sTagSection = cCtrL.Tag
    
    vTagSectionParams = Split(sTagSection, ".")
    sParam = "TrgtMSelect"
    iWhere = InStr(1, vTagSectionParams(0), sParam)
    
    'Se houver na TAG do controle a indicação [ TrgtMSelect ]
    If iWhere > 0 Then
    
        'Verifica se a TAG informa o [ sFilGrp ] contendo o [ Grupo de Filtragem ] que deverá ser manipulado
        sParam = "Grp"
            'sStR1 = "Formulário:  [ " & sForM & " ]" & vbCr & "TargetCtrl: " & "   [ " & sTrgtCtrL & " ]" & vbCr & "-------------------------------------------------------------------------------"
            'sStR2 = " O [ grupo de filtragem ] do TargetCtrl não foi informado" & vbCr & "  e ele não poderá ser pesquisado."
            
            'sFilGrp = GetTagParams()
            If sFilGrp = "" Then Exit Sub
        
        'Identifica o [ TrgtCtrl ] de [ sFilGrp ]
        If IsObject(dictFormFilterGrpsTrgts(sForM)(sFilGrp)) Then
            Set clObjTargtCtrlParam = dictFormFilterGrpsTrgts(sForM)(sFilGrp)
            vA = clObjTargtCtrlParam.sTargtCtrlName
        
        End If
Stop
        
      
        
    End If


End Sub
