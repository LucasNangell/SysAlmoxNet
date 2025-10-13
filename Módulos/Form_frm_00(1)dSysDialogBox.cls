VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_00(1)dSysDialogBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnClose_Click()

    DoCmd.Close acForm, Me.Name, acSaveYes
    
End Sub

Private Sub btnNo_Click()
    TempVars("SysDialogBox") = False
    DoCmd.Close acForm, Me.Name, acSaveYes
    
End Sub

Private Sub btnYes_Click()
    TempVars("SysDialogBox") = True
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub Form_Load()
    lblTitulo.Caption = Split(Me.OpenArgs, "|")(0)
    lblTexto.Caption = Split(Me.OpenArgs, "|")(1)
    '------------------------------------------------------------
    'Define posição do formulário centralizado sobre o formulário que acionou o [ SysDialogBox ]
    '------------------------------------------------------------
    
    clObjFormOpenParams.bSetPosition = True
    clObjFormOpenParams.bCentralizeForm = True
    clObjFormOpenParams.sTrgtForm = Me.Name
    Set clObjFormOpenParams.fTrgtForm = Me
    
    Call FormLoad08_SetFormPosition(clObjFormOpenParams)
    
End Sub
