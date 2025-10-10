VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90aDevButtons-All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btnMinimizar_Click()
    
    'Reduzir a janela do Access e posicionar pro centro da tela, pra ficar por trás do Form
    DoCmd.RunCommand acCmdAppRestore
    Call AccessMoveSize(500, 300, 500, 300)   'Left, Top, Width, Heigth

End Sub


Private Sub btSair_Click()
    
    DoCmd.Close acForm, Me.Name, acSavePrompt
    'DoCmd.Quit
    
End Sub


Private Sub btnSair_Click()
    Dim vA, vB
    Dim lgInT As Integer
    Dim DbO As Object
    Dim sOpenedForm As String
    Dim bFrmFound As Boolean
    Dim sFrmToFind As String
    Dim aObJ As AccessObject, oDbS As Object
    Dim iFrmIndexID As Integer
    
    vA = Me.Name

    DoCmd.Close , , acSavePrompt
'Stop
    lgInT = Forms.Count
    sFrmToFind = "frm01aOPsSearch"
    
    For Each oDbS In Forms
        sOpenedForm = oDbS.Name
        If sOpenedForm = sFrmToFind Then bFrmFound = True
    Next oDbS
    
    If bFrmFound Then
Stop
        iFrmIndexID = dictFormsSeqID(Forms(sFrmToFind).Name)
        Call EnableCtrl(iFrmIndexID, Forms(sFrmToFind)!btnSair, True)

    End If
    
    If sOpenedForm = "" Then
        dictFormsSeqID.RemoveAll
        'esvazia o dicionário retirando todos os forms encontrados no sistema.
        'do contrário a cada vez que o sistema é fechado, na primeira vez em que é iniciado novamente
        ' causa um erro na linha que cria o dicionário do zero, pois o dicionário já existe
    
    End If
End Sub
