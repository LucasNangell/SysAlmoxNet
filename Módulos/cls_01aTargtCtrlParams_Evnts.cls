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
Public sClsLstbxSQL_dGROUPBY As String
Public sClsLstbxSQL_eHAVING As String
Public sClsLstbxSQL_fOrderBy As String  'SQL original do ListBox ou Combo ao carregar o Form (OrderBy)
Public sClsLstbxSQL_gMAIN As String     'SQL original do ListBox ou Combo ao carregar o Form (Select + From)

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
    cCtrL.BeforeUpdate = "[Event Procedure]"
    cCtrL.AfterUpdate = "[Event Procedure]"
    cCtrL.OnMouseDown = "[Event Procedure]"

End Property

Private Sub mListBox_BeforeUpdate(Cancel As Integer)

    Dim vA, vC, vD
    Dim vB
    Dim cCtrL As Control
    Dim sCtrL As String
    Dim fForM As Form
    Dim sFilGrp As String
    Dim sForM As String

    
    Set cCtrL = mListBox
    Set fForM = cCtrL.Parent
    sForM = fForM.Name
    sCtrL = cCtrL.Name

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ] LostFocus, controle [ " & sCtrL & "]   "
'Stop
    
'   'Chama a rotina de filtragem
    Call pb_TargtCtrlUpdate03_UNIQUEupdate(fForM, cCtrL)
    Call PbSubDataFields_FillFromListbox(cCtrL)
                                
End Sub


Private Sub mListBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim vA, vB
    Dim vKeyGrp, vKeyCtrl
    Dim sTrgtCtrlFilGrp As String, sForM As String
    Dim sTrggCtrlFilGrp As String
    Dim cEdtCtrl As Control
    Dim fForM As Form
    Dim rsList As Recordset
    Dim fField As Field
    
    '---------------------------------------------------------------
    ' chama pbSub00_CtrlsEnblDsble_GetParams
    ' pra habilitar/desabilitar os controles depois que um valor é selecionado na Listbox
    '---------------------------------------------------------------
'Stop
    Set fForM = mListBox.Parent
'MsgBox "Mouse Down [ " & mListBox.Name & " ]"
'Stop
    'Quando um item é selecionado na listbox
    ' habilita o botão de Edição de dados do grupo de filtragem indicado
    ' se a lista estiver liberada pra seleção dos valores
    ' Caso contrário nenhuma ação deve ser executada
    If mListBox.Locked = False Then

        sForM = mListBox.Parent.Name
        
        'Verifica se o listbox é um [ TrggCtrl ], caso seja
        ' adiciona os valores dos campos do item selecionado ao [ dictFieldsSelectedItem ]
        If IsObject(dictTrggCtrlsInForm(sForM)(mListBox.Name)) Then
            sTrggCtrlFilGrp = dictTrggCtrlsInForm(sForM)(mListBox.Name).sFilGrpToFilter
            
            If Not IsObject(dictFieldsSelectedItem(sForM)) Then Set dictFieldsSelectedItem(sForM) = New Dictionary
            If Not IsObject(dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp)) Then Set dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp) = New Dictionary
            dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp).RemoveAll
            
            Set rsList = CurrentDb.OpenRecordset(mListBox.RowSource)
            
            For Each fField In rsList.Fields
                dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp).Add fField.Name, fField.Value
            Next fField
            
            rsList.Close
            Set rsList = Nothing
            
        End If

        Call pbSub00_CtrlsEnblDsble_GetParams(fForM, mListBox.Name & "Select")
        
        
    End If
End Sub

