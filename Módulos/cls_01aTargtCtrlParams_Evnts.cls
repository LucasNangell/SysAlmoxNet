VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_01aTargtCtrlParams_Evnts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Public cCtrl As Control          'Controle
Public sTargtCtrlName As String  'Nome do controle
                                ' nome do controle, usado como Indexador do controle dentro do dicionário
                                ' meramente informativo
                            
Public sFilGrp As String    'Grupo de filtragem a que pertence o controle,
                            ' neste Objeto de Classe é meramente informativo
                            
'Public dictFormFilterGrpTrgts As New Dictionary  'Adiciona no dicionário interno de cada Grupo de Filtragem, os [ TrgtCtrls ] associados ao grupo


Public sRecCntCtrlName As String 'Guarda o nome do controle que deve exibir a contagem de registros do grupo Trgt Ctrl
'Public cRecCntCtrl As Control  'Controle no formulário que exibe a contagem de registros da Listbox

Public sClsLstbxSQL_aSELECT As String   'SQL original do ListBox ou Combo ao carregar o Form (Select)
Public sClsLstbxSQL_bFROM As String     'SQL original do ListBox ou Combo ao carregar o Form (From)
Public sClsLstbxSQL_cWHERE As String    'SQL original do ListBox ou Combo ao carregar o Form (Where)
Public sClsLstbxSQL_dOrderBy As String  'SQL original do ListBox ou Combo ao carregar o Form (OrderBy)

Public sClsLstbxSQL_eMAIN As String     'SQL original do ListBox ou Combo ao carregar o Form (Select + From)

Public dictTrgtQryFields As New Dictionary

'Public sFilGrp As Integer    ' grupo de filtragem a que pertence o controle, o que permite identificar
                                 ' qual controle deve exibir os resultados da pesquisa e se há outros controles
                                 ' a serem considerados
                                 
                                 'essa informação foi transferida como identificador do ListBox detro do dicionário
                                 ' e não precisa mais ser usada como um parâmetro da Classe

'parei aqui: carregar o parâmetro no dicionário
Public bMltSlct As Boolean          'indica se o Listbox deve ser Single ou Multiselect


Private WithEvents mListBox As ListBox
Attribute mListBox.VB_VarHelpID = -1


'------------------------------------------------------------
'Trecho pra manipular eventos dos controles
'----------------------------------
Public Property Set InitCtrl(cCtrL As Control)
    Dim vA, vB

    Set mListBox = Nothing
    Set mListBox = cCtrL
    cCtrL.AfterUpdate = "[Event Procedure]"
    cCtrL.OnMouseDown = "[Event Procedure]"

End Property

Private Sub mListBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim vA, vB
    Dim vKeyGrp, vKeyCtrl
    Dim sFilGrp As String, sForM As String
    Dim cEdtCtrl As Control
    Dim fForM As Form
    
    '---------------------------------------------------------------
    ' chama pbSub00_CtrlsEnblDsble_GetParams
    ' pra habilitar/desabilitar os controles depois que um valor é selecionado na Listbox
    '---------------------------------------------------------------
    
    Set fForM = mListBox.Parent
Stop
    'Quando um item é selecionado na listbox
    ' habilita o botão de Edição de dados do grupo de filtragem indicado
    ' se a lista estiver liberada pra seleção dos valores
    ' Caso contrário nenhuma ação deve ser executada
    If mListBox.Locked = False Then

        sForM = mListBox.Parent.Name

        'Encontra o grupo de filtragem da lista [ TrgtCtrl ]
        ' verifica se a lista existe no dict [ dictTrgtCtrlsFilterGrps(sForM) ] pra buscar o grupo diretamente
        If dictTrgtCtrlsFilterGrps(sForM).Exists(mListBox.Name) Then
            sFilGrp = dictTrgtCtrlsFilterGrps(sForM)(mListBox.Name)
        
        Else
        
        'Caso a [ mListBox ] não exista em [ dictTrgtCtrlsFilterGrps(sForM) ],
        ' percorre cada item do [ dictFormFilterGrpTrgts(sForM)] para encontrar o grupo da lista
            For Each vKeyGrp In dictFormFilterGrpTrgts(sForM)
                Set clObjTargtCtrlParam = dictFormFilterGrpTrgts(sForM)(vKeyGrp)
                If clObjTargtCtrlParam.sTargtCtrlName = mListBox.Name Then sFilGrp = vKeyGrp
            
            Next vKeyGrp
        
        End If

        
        Call pbSub00_CtrlsEnblDsble_GetParams(fForM, "RecSelect", sFilGrp)
        
        'Identifica o botão de edição do grupo de filtragem pra habilitá-lo,
        ' por meio do parâmetro ActType> n TAG do botão
        ' se for confirmada a existência da lista em [ dictTrgtCtrlsFilterGrps(sForM) ]
        ' encontra o botão [ EditRec ] referente ao [ sFilGrp ] da [ mListBox ]
'        For Each vKeyCtrl In dictFormCommButtons(sForM)
'            Set clObjCommButtons = dictFormCommButtons(sForM)(vKeyCtrl)
'            Select Case clObjCommButtons.sActType
'
'                Case "EditRec", "DupRec", "ExcRec"
'                    Set cEdtCtrl = Forms(sForM).Controls(clObjCommButtons.sCtrlName)
'                    If cEdtCtrl.Enabled = False And clObjCommButtons.sFilGrp = sFilGrp Then
'                        'Chama [ pbSub02_CtrlsEnblDsbl_Apply ] pra habilitar o botão específico sem precisar percorrer por todos os controles
'                        Call pbSub02_CtrlsEnblDsbl_Apply(mListBox.Parent, cEdtCtrl, "btn", True, True)
'                    End If
'                End Select
'
'        Next vKeyCtrl

    End If
End Sub

Private Sub mListBox_AfterUpdate()
Stop
    
    Call PbSubFillFieldsByList(mListBox)

End Sub
