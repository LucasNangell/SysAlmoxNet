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

Private Sub btnCalendar1_Click()
    Call CalendarButton(Me.Name, ActiveControl, Me.txtStkDataEntrada, "Selecione uma data de entrada")

End Sub

Private Sub btnCalendar2_Click()
    Call CalendarButton(Me.Name, ActiveControl, Me.txtStkValidade, "Selecione uma data de validade")
    
End Sub


Private Sub btnClose_Click()

    DoCmd.Close

End Sub


Private Sub btnEditarEntrada_Click()

    txtStkDataEntrada.SetFocus
    
End Sub


Private Sub btnReset2Search_Click()
    
    Me.txtSrchDescProd.SetFocus

End Sub

Private Sub Comando545_Click()

End Sub

Private Sub Form_Load()

'Stop
    
    If Not clObjFormOpenParams Is Nothing Then
        Set clObjFormOpenParams.fTrgtForm = Me
'Stop
        Call FormLoad06a_BackFromFormLoad(clObjFormOpenParams)
    End If

    'Me.lstEstoqueIn.RowSource = ""
    'Me.lstEstoqueSaldo.RowSource = ""
    Me.lstProdsBase.Selected(0) = True
    
    
End Sub


Private Sub lstProdsBase_AfterUpdate()
    
'    'Seleciona o primeiro item da lista
'    Me.lstEstoqueIn.Selected(0) = True

End Sub

Private Sub txtPrecoEmb_Change()
    Call MskdTxtbox02_TextMask(ActiveControl, ActiveControl.Text, "R$ #,###.00", 6)
    'Call MskdTxtbox02_TextMask(ActiveControl, ActiveControl.Text, "P_###,###", 6)
    
End Sub
