VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_02(1)aProdCadastro - bk"
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

