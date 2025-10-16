VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_01(1)cProdEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCalendar_Click()
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    clObjFormOpenParams.sTrggForM = Me.Name
    clObjFormOpenParams.lngFormLeft = btnCalendar.Left
    clObjFormOpenParams.lngFormTop = btnCalendar.Top
    
    Call CalendarFor(Me.txtStkDataEntrada, "Selecione uma data de entrada")
    Set clObjFormOpenParams = Nothing

End Sub

Private Sub btnCancelarEntrada_Click()

   Dim vA, vB, vC
   Dim sTitulo As String, sTexto As String, sForM As String
   sForM = Me.Name

'Stop
    
    sTitulo = "Cancelamento de Edição de Dados"
    sTexto = "As alterações realizadas serão perdidas." & vbCrLf & "Confirma?"
    TempVars("bReply") = DialogBoxReply(sTitulo, sTexto, sForM)
'Stop

End Sub

Private Sub btnClose_Click()

    DoCmd.Close

End Sub


Private Sub btnEditarEntrada_Click()

    txtStkDataEntrada.SetFocus
    
End Sub

Private Sub btnExcluirEntrada_Click()
   Dim sTitulo As String, sTexto As String, sForM As String
   sForM = Me.Name

    sTitulo = "Exclusão de Entrada no estoque"
    sTexto = "A entrada selecionada será excluída permanentemente." & vbCrLf & "Confirma?"
    TempVars("bReply") = DialogBoxReply(sTitulo, sTexto, sForM)

End Sub

Private Sub btnExcluirPrd_Click()
   Dim sTitulo As String, sTexto As String, sForM As String
   sForM = Me.Name

    sTitulo = "Exclusão de Produto"
    sTexto = "O produto selecionado será excluído permanentemente." & vbCrLf & "Confirma?"
    TempVars("bReply") = DialogBoxReply(sTitulo, sTexto, sForM)
End Sub

Private Sub btnReset2Search_Click()
    
    Me.txtSrchDescProd.SetFocus

'Stop
    

End Sub

Private Sub Form_Load()

'Stop
    
    If Not clObjFormOpenParams Is Nothing Then
        Set clObjFormOpenParams.fTrgtForm = Me
'Stop
        Call FormLoad06a_BackFromFormLoad(clObjFormOpenParams)
    End If

    Me.lstEstoqueIn.RowSource = ""
    Me.lstEstoqueSaldo.RowSource = ""
    
End Sub


