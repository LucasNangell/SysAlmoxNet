VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90eWidthColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim fTrgtForm As Form
Dim cTrgtCtrl As Control
Dim sTrgtCtrl As String
Dim cHeadCtrl As Control
Dim qDef As QueryDef
Dim fField As Field
Dim bSelectedList As Boolean
Dim bSelectingHeads As Boolean
Dim iSelectedItem As Integer
Dim iOldSelectedItem As Integer
Dim sOldWidthColum As String
Dim sHead As String, sOldHead As String
Dim dictQryListFields As New Dictionary
Private Sub btnAddColumn_Click()
    lstSelectedColumns.AddItem lstSelectedColumns.ListCount + 1 & ";" & lstFields.Column(0, lstFields.ListIndex) & ";;;" & lstFields.ListIndex
    
End Sub

Private Sub btnCancelar_Click()
    'DoCmd.Close acForm, fTrgtForm.Name, acSaveNo
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub btnFinalizar_Click()
    bSelectingHeads = True
    bSelectedList = False
End Sub

Private Sub btnSalvar_Click()
    Dim sWidthColumns As String
    Dim iInT As Integer
    Dim iInT2 As Integer
    Dim bBoL As Boolean
    
    Set cTrgtCtrl = fTrgtForm.Controls(txtSelectedList)
    sWidthColumns = ""
    
    For iInT2 = 1 To dictQryListFields.Count
        bBoL = False
        For iInT = 0 To lstSelectedColumns.ListCount - 1
            If lstSelectedColumns.Column(1, iInT) = dictQryListFields(iInT2) Then
                sWidthColumns = sWidthColumns & lstSelectedColumns.Column(3, iInT) & ";"
                bBoL = True
            End If
        Next iInT
        If Not bBoL Then
            sWidthColumns = sWidthColumns & "0;"
        End If
    Next iInT2
    cTrgtCtrl.ColumnCount = lstFields.ListCount
    cTrgtCtrl.ColumnWidths = sWidthColumns
    
    'DoCmd.Save acForm, fTrgtForm.Name
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()
Dim fForm As Form

If Forms.Count > 2 Then
    MsgBox "Por favor deixe apenas o formulário que terá as listas manipuladas aberto antes de abrir este formulário.", vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End If

bSelectedList = False
bSelectingHeads = False

For Each fForm In Forms
    If fForm.Name <> Me.Name Then Set fTrgtForm = fForm
Next fForm
End Sub

Private Sub Form_Timer()
On Error Resume Next

If IsNull(txtSelectedList) Then bSelectedList = False

If sTrgtCtrl <> fTrgtForm.ActiveControl.Name Then
    Set cTrgtCtrl = fTrgtForm.ActiveControl
    sTrgtCtrl = cTrgtCtrl.Name
    If cTrgtCtrl.ControlType = acListBox Then bSelectedList = False
    
    If bSelectedList = False And Not bSelectingHeads Then
        If cTrgtCtrl.ControlType = acListBox Then
            dictQryListFields.RemoveAll
            lstSelectedColumns.RowSource = ""
            bSelectedList = True
            txtSelectedList = cTrgtCtrl.Name
            lstFields.RowSource = ""
            Set qDef = CurrentDb.QueryDefs(cTrgtCtrl.RowSource)
            For Each fField In qDef.Fields
                lstFields.AddItem fField.Name, lstFields.ListCount
                dictQryListFields.Add lstFields.ListCount, fField.Name
            Next fField
        Else
            txtSelectedList = "Selecione alguma List box"
        End If
        
    End If
    iSelectedItem = lstSelectedColumns.ListIndex
    If iOldSelectedItem <> iSelectedItem Then Call lstSelectedColumns_AfterUpdate
    iOldSelectedItem = iSelectedItem
    
End If
'If bSelectingHeads Then Set cHeadCtrl = cTrgtCtrl

End Sub

Private Sub lstSelectedColumns_AfterUpdate()
Dim vDadosLst As Variant
Dim vA
Dim iInT As Integer, iInT2 As Integer, iConT As Integer
Dim sNewRowSource As String
 iSelectedItem = lstSelectedColumns.ListIndex
    If cTrgtCtrl.ControlType = acCommandButton Then
        sHead = cTrgtCtrl.Caption
        If sHead <> sOldHead Then
            sOldHead = sHead
            If iSelectedItem > -1 Then
                ReDim vDadosLst(1 To lstSelectedColumns.ListCount, 1 To lstSelectedColumns.ColumnCount)
                iConT = 0
                vA = Split(lstSelectedColumns.RowSource, ";")
                For iInT = 1 To lstSelectedColumns.ListCount
                    For iInT2 = 1 To lstSelectedColumns.ColumnCount
                        vDadosLst(iInT, iInT2) = vA(iConT)
                        iConT = iConT + 1
                    Next iInT2
                Next iInT
                sNewRowSource = ""
                vDadosLst(lstSelectedColumns.ListIndex + 1, 3) = cTrgtCtrl.Caption
                vDadosLst(lstSelectedColumns.ListIndex + 1, 4) = cTrgtCtrl.Width
                For iInT = 1 To lstSelectedColumns.ListCount
                    For iInT2 = 1 To lstSelectedColumns.ColumnCount
                        sNewRowSource = sNewRowSource & vDadosLst(iInT, iInT2) & ";"
                    Next iInT2
                Next iInT
                
                sNewRowSource = Left(sNewRowSource, Len(sNewRowSource) - 1)
                
                lstSelectedColumns.RowSource = sNewRowSource
    
            Else
            
            End If
        End If
    End If
End Sub

Private Sub lstSelectedColumns_BeforeUpdate(Cancel As Integer)

End Sub
