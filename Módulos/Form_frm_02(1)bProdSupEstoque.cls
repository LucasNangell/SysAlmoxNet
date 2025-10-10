VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_02(1)bProdSupEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub lstProdsBase_AfterUpdate()

    Dim vA, vB
    Dim iListIndex As Integer
    Dim vListboxQryValue As Variant
    
    
'Stop
    'Identifica o registro selecionado na Listbox
    iListIndex = Me!lstProdsBase.ListIndex
    
    'Exibe no [ txtStkProdDescriç ] o valor da 2a coluna de dados da consulta fonte de dados
    vListboxQryValue = Me!lstProdsBase.Column(1, iListIndex)
    txtStkProdDescriç = vListboxQryValue
    
    'Exibe nos controles do box de edição do produto os valores de cada um dos campos que formam o produto
    txtEdtProd = Me!lstProdsBase.Column(2, iListIndex)
    txtEdtVariaç = Me!lstProdsBase.Column(3, iListIndex)
    txtEdtCor = Me!lstProdsBase.Column(5, iListIndex)
    txtEdtMaterial = Me!lstProdsBase.Column(7, iListIndex)
    txtEdtMedida = Me!lstProdsBase.Column(9, iListIndex)
    txtEdtComplem = Me!lstProdsBase.Column(10, iListIndex)
    
'Stop
    txtEdtUnProd = Me!lstProdsBase.Column(16, iListIndex)
    txtEdtUnConsumo = Me!lstProdsBase.Column(18, iListIndex)
    txtEdtUnPed = Me!lstProdsBase.Column(20, iListIndex)
    txtEdtQtdMin = Me!lstProdsBase.Column(21, iListIndex)
    
    'cmbEdtCateg
    'cmbEdtAplic
    'cmbEdtSetor
    
    
End Sub


Private Sub txtSrchVariaç_Change()
    Call pb_TargtCtrlUpdate00_TimerDelay
End Sub
