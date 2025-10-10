VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_01(1)cProdEstoque - bk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnClose_Click()

    DoCmd.Close

End Sub


Private Sub Form_Load()
'Stop
    If Not clObjFormOpenParams Is Nothing Then
        Set clObjFormOpenParams.fTrgtForm = Me
        Call FormLoad06a_BackFromFormLoad(clObjFormOpenParams)
    End If
End Sub


Private Sub lstEstoqueIn_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    'Call pbSub00_CtrlsEnblDsble_GetParams(Me, "EntradaSelect")
    
End Sub


Private Sub lstProdsBase_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    'Call pbSub00_CtrlsEnblDsble_GetParams(Me, "RecSelect")

End Sub
