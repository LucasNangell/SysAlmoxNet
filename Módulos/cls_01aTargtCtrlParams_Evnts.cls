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
    Dim sCtrl As String
    Dim fForM As Form
    Dim sFilGrp As String
    Dim sForM As String
    Dim sTrggCtrlFilGrp As String
    Dim rsList As Recordset
    Dim fField As Field
    Dim iQryID As Integer
    Dim sQryIDName As String
    Dim qDef As QueryDef
    Dim sSQL As String
    Dim sWhere As String
    Dim sLastPartSQL As String
    
    Set cCtrL = mListBox
    Set fForM = cCtrL.Parent
    sForM = fForM.Name
    sCtrl = cCtrL.Name

'MsgBox "teste --------------------------------------------------------------------------" & vbCr & "Aciona [ HighlightClrChange ] LostFocus, controle [ " & sCtrL & "]   "
'Stop
    If mListBox.Locked = False Then Call pbSub00_CtrlsEnblDsble_GetParams(fForM, mListBox.Name & "Select")
    'Verifica se o listbox é um [ TrggCtrl ], caso seja
    ' adiciona os valores dos campos do item selecionado ao [ dictFieldsSelectedItem ]
    If IsObject(dictTrggCtrlsInForm(sForM)(mListBox.Name)) Then
        sTrggCtrlFilGrp = dictTrggCtrlsInForm(sForM)(mListBox.Name).sFilGrpToFilter
        
        If Not IsObject(dictFieldsSelectedItem(sForM)) Then Set dictFieldsSelectedItem(sForM) = New Dictionary
        If Not IsObject(dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp)) Then Set dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp) = New Dictionary
        dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp).RemoveAll
        
        If InStr(mListBox.RowSource, "SELECT") > 0 Then
            sSQL = mListBox.RowSource
        Else
            sSQL = Replace(CurrentDb.QueryDefs(mListBox.RowSource).sql, ";", "")
        End If
        'sSQL = IIf(InStr(mListBox.RowSource, "SELECT") > 0, mListBox.RowSource, CurrentDb.QueryDefs(mListBox.RowSource).sql)
        If InStr(sSQL, "GROUP BY") > 0 Then
            sLastPartSQL = "GROUP BY " & Split(sSQL, "GROUP BY")(1)
            sSQL = Split(sSQL, "GROUP BY")(0)
        ElseIf InStr(sSQL, "ORDER BY") > 0 Then
            sLastPartSQL = "ORDER BY " & Split(sSQL, "ORDER BY")(1)
            sSQL = Split(sSQL, "ORDER BY")(0)
        End If
                
        Set rsList = CurrentDb.OpenRecordset(mListBox.RowSource)
        
        sQryIDName = rsList.Fields(0).Name
        iQryID = mListBox.Column(0, mListBox.ListIndex)
            
        sWhere = "WHERE ([" & sQryIDName & "] = " & iQryID & ")"
        
        sSQL = sSQL & sWhere & vbCr & sLastPartSQL
        
        Set rsList = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)
        
        If rsList.RecordCount = 1 Then
            For Each fField In rsList.Fields
                dictFieldsSelectedItem(sForM)(sTrggCtrlFilGrp).Add fField.Name, fField.Value
            Next fField
        End If
        
        rsList.Close
        Set rsList = Nothing
    
    Call pb_TargtCtrlUpdate03_UNIQUEupdate(fForM, cCtrL)
    End If
    'Chama a rotina de filtragem
    
    Call PbSubDataFields_FillFromListbox(cCtrL)
                                
    'Libera atualização de tela

End Sub



    



