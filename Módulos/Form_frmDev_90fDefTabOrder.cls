VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90fDefTabOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim fTrgtForm As Form
Dim cTrgtCtrl As Control
Dim sTrgtCtrl As String

Dim dictTabIndexCtrls As New Dictionary




Private Sub btnCancelar_Click()
    DoCmd.Close acForm, Me.Name
    
End Sub



Private Sub btnLimparList_Click()
    lstCtrls.RowSource = ""
End Sub

Private Sub btnMoveDown_Click()
Dim dictTemp As New Dictionary
Dim vKeyIndex As Variant
Dim iLstIndex As Integer
Dim iNewIndex As Integer

dictTemp.RemoveAll
iLstIndex = lstCtrls.ListIndex

If iLstIndex > -1 And iLstIndex < lstCtrls.ListCount - 1 Then
    For Each vKeyIndex In dictTabIndexCtrls
        dictTemp.Add vKeyIndex, dictTabIndexCtrls(vKeyIndex)
    Next vKeyIndex
    dictTabIndexCtrls.RemoveAll
    For Each vKeyIndex In dictTemp
        If Not dictTabIndexCtrls.Exists(vKeyIndex) Then
            If vKeyIndex = iLstIndex + 1 Then
                dictTabIndexCtrls.Add iLstIndex + 1, dictTemp(iLstIndex)
            ElseIf vKeyIndex > iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex)
            ElseIf vKeyIndex = iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex + 1)
            ElseIf vKeyIndex < (iLstIndex + 1) Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex)
            End If
        End If
    Next vKeyIndex
    dictTemp.RemoveAll
    lstCtrls.RowSource = ""
    For Each vKeyIndex In dictTabIndexCtrls
        lstCtrls.AddItem vKeyIndex & ";" & dictTabIndexCtrls(vKeyIndex).Name, lstCtrls.ListCount
    Next vKeyIndex
    lstCtrls.Selected(iLstIndex + 1) = True
End If

End Sub

Private Sub btnMoveUp_Click()
Dim dictTemp As New Dictionary
Dim vKeyIndex As Variant
Dim iLstIndex As Integer
Dim iNewIndex As Integer

dictTemp.RemoveAll

iLstIndex = lstCtrls.ListIndex

If iLstIndex > 0 Then
    For Each vKeyIndex In dictTabIndexCtrls
        dictTemp.Add vKeyIndex, dictTabIndexCtrls(vKeyIndex)
    Next vKeyIndex
    dictTabIndexCtrls.RemoveAll
    For Each vKeyIndex In dictTemp
        If Not dictTabIndexCtrls.Exists(vKeyIndex) Then
            If vKeyIndex < (iLstIndex - 1) Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex)
            ElseIf vKeyIndex = iLstIndex - 1 Then
                dictTabIndexCtrls.Add iLstIndex - 1, dictTemp(iLstIndex)
            ElseIf vKeyIndex = iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex - 1)
            ElseIf vKeyIndex > iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex)
            End If
        End If
    Next vKeyIndex
    dictTemp.RemoveAll
    lstCtrls.RowSource = ""
    For Each vKeyIndex In dictTabIndexCtrls
        lstCtrls.AddItem vKeyIndex & ";" & dictTabIndexCtrls(vKeyIndex).Name, lstCtrls.ListCount
    Next vKeyIndex
    lstCtrls.Selected(iLstIndex - 1) = True
End If

End Sub


Private Sub btnSalvar_Click()
    Dim iInT As Integer
    Dim iIndeX As Integer
    Dim bBoL As Boolean
    Dim vKeyIndex As Variant
    
    
    For Each vKeyIndex In dictTabIndexCtrls
        Set cCtrL = dictTabIndexCtrls(vKeyIndex)
        cCtrL.TabIndex = vKeyIndex
    Next vKeyIndex

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()
Dim fForm As Form

If Forms.Count > 2 Then
    MsgBox "Por favor deixe apenas o formulário que terá as listas manipuladas aberto antes de abrir este formulário.", vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End If
dictTabIndexCtrls.RemoveAll

For Each fForm In Forms
    If fForm.Name <> Me.Name Then Set fTrgtForm = fForm
Next fForm
End Sub

Private Sub Form_Timer()
    On Error Resume Next
    DoEvents

    

    Set cTrgtCtrl = fTrgtForm.ActiveControl
    If sTrgtCtrl <> cTrgtCtrl.Name Then
    
        sTrgtCtrl = cTrgtCtrl.Name
        
        If Not dictTabIndexCtrls.Exists(sTrgtCtrl) And sTrgtCtrl <> "" Then
            dictTabIndexCtrls.Add lstCtrls.ListCount, cTrgtCtrl
            lstCtrls.AddItem lstCtrls.ListCount & ";" & sTrgtCtrl, lstCtrls.ListCount
            If lstCtrls.ListIndex > -1 Then lstCtrls.Selected(lstCtrls.ListIndex) = False
        Else
        
        End If
        
        DoCmd.SelectObject acForm, fTrgtForm.Name, False
        DoCmd.RunCommand acCmdSelectForm
    End If
    
End Sub

Private Sub lstCtrls_DblClick(Cancel As Integer)
    Dim iLstIndex As Integer
    Dim vKeyIndex As Variant
    Dim iInT As Integer
    Dim dictTemp As New Dictionary
    iLstIndex = lstCtrls.ListIndex
    If iLstIndex < 0 Then Exit Sub
    For iInT = 0 To lstCtrls.ListCount - 1
        dictTemp.Add iInT, fTrgtForm.Controls(lstCtrls.Column(1, iInT))
    Next iInT
    dictTabIndexCtrls.RemoveAll
    lstCtrls.RowSource = ""
    For Each vKeyIndex In dictTemp
        If Not dictTabIndexCtrls.Exists(vKeyIndex) Then
            If vKeyIndex < iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex, dictTemp(vKeyIndex)
            ElseIf vKeyIndex > iLstIndex Then
                dictTabIndexCtrls.Add vKeyIndex - 1, dictTemp(vKeyIndex)
            End If
        End If
    Next vKeyIndex
    For Each vKeyIndex In dictTabIndexCtrls
        lstCtrls.AddItem vKeyIndex & ";" & dictTabIndexCtrls(vKeyIndex).Name, lstCtrls.ListCount
    Next vKeyIndex
    
End Sub
