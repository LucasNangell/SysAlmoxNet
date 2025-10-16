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
Dim dictSelectedFields As New Dictionary
Dim bAjustando As Boolean
Dim bChanged As Boolean

Private Sub btnAdd_Click()
    Dim lngWidth As Long
    Dim iInT As Integer, iInT2 As Integer
    Dim iCoL As Integer
    Dim iListIndex As Integer
    Dim sNewItem As String, sWidthColumns As String
    Dim vA
    Dim bBoL As Boolean
    iListIndex = lstSelectedColumns.ListIndex
    If iListIndex > -1 Then
        
        lngWidth = txtWidth
        lngWidth = lngWidth + 5
        txtWidth = lngWidth
        
        ReDim vA(lstSelectedColumns.ListCount - 1, lstSelectedColumns.ColumnCount - 1)
        
        For iInT = 0 To lstSelectedColumns.ListCount - 1
            For iCoL = 0 To lstSelectedColumns.ColumnCount - 1
                vA(iInT, iCoL) = lstSelectedColumns.Column(iCoL, iInT)
            Next iCoL
        Next iInT
        
        vA(iListIndex, 3) = lngWidth
        
        lstSelectedColumns.RowSource = ""
        
        For iInT = 0 To UBound(vA)
        sNewItem = ""
            For iCoL = 0 To lstSelectedColumns.ColumnCount - 1
                sNewItem = sNewItem & vA(iInT, iCoL) & ";"
            Next iCoL
            lstSelectedColumns.AddItem sNewItem, lstSelectedColumns.ListCount
        Next iInT
        
        Set cTrgtCtrl = fTrgtForm.Controls(txtSelectedList)
        sWidthColumns = ""
    
        For iInT2 = 1 To dictQryListFields.Count
            bBoL = False
            For iInT = 0 To lstSelectedColumns.ListCount - 1
                If lstSelectedColumns.Column(3, iInT) = "" Then
                    MsgBox "Existem colunas sem HeadCtrl selecionado, selecione todos os HeadCtrls antes de prosseguir.", vbInformation
                    Exit Sub
                End If
                If lstSelectedColumns.Column(1, iInT) = dictQryListFields(iInT2) Then
                    sWidthColumns = sWidthColumns & lstSelectedColumns.Column(3, iInT) & ";"
                    bBoL = True
                End If
            Next iInT
            If Not bBoL Then
                sWidthColumns = sWidthColumns & "0;"
            End If
        Next iInT2

    cTrgtCtrl.ColumnWidths = sWidthColumns
    
    End If
    
    fTrgtForm.Repaint
    
End Sub

Private Sub btnAddColumn_Click()
    
    Dim iInT As Integer
    Dim vA
    If lstFields.ListIndex > -1 Then
        For iInT = 1 To dictQryListFields.Count
            If dictQryListFields(iInT) = lstFields.Column(0, lstFields.ListIndex) Then
                If Not IsEmpty(dictQryListFields(iInT)) Then dictSelectedFields.Add iInT, dictQryListFields(iInT)
            End If
        Next iInT
        lstSelectedColumns.RowSource = ""
        For Each vA In dictSelectedFields
            lstSelectedColumns.AddItem lstSelectedColumns.ListCount + 1 & ";" & dictSelectedFields(vA) & ";;;" & vA
        Next vA
        lstFields.RemoveItem (lstFields.ListIndex)
    End If
End Sub

Private Sub btnCancelar_Click()
    bAjustando = False
    DoCmd.OpenForm fTrgtForm.Name, acDesign
    DoCmd.Close acForm, Me.Name, acSaveYes
    
End Sub


Private Sub btnMin_Click()
    Dim lngWidth As Long
    Dim iInT As Integer, iInT2 As Integer
    Dim iCoL As Integer
    Dim iListIndex As Integer
    Dim sNewItem As String, sWidthColumns As String
    Dim vA
    Dim bBoL As Boolean
    iListIndex = lstSelectedColumns.ListIndex
    If iListIndex > -1 Then
        
        lngWidth = txtWidth
        lngWidth = lngWidth - 5
        txtWidth = lngWidth
        
        ReDim vA(lstSelectedColumns.ListCount - 1, lstSelectedColumns.ColumnCount - 1)
        
        For iInT = 0 To lstSelectedColumns.ListCount - 1
            For iCoL = 0 To lstSelectedColumns.ColumnCount - 1
                vA(iInT, iCoL) = lstSelectedColumns.Column(iCoL, iInT)
            Next iCoL
        Next iInT
        
        vA(iListIndex, 3) = lngWidth
        
        lstSelectedColumns.RowSource = ""
        
        For iInT = 0 To UBound(vA)
        sNewItem = ""
            For iCoL = 0 To lstSelectedColumns.ColumnCount - 1
                sNewItem = sNewItem & vA(iInT, iCoL) & ";"
            Next iCoL
            lstSelectedColumns.AddItem sNewItem, lstSelectedColumns.ListCount
        Next iInT
        
        Set cTrgtCtrl = fTrgtForm.Controls(txtSelectedList)
        sWidthColumns = ""
    
        For iInT2 = 1 To dictQryListFields.Count
            bBoL = False
            For iInT = 0 To lstSelectedColumns.ListCount - 1
                If lstSelectedColumns.Column(3, iInT) = "" Then
                    MsgBox "Existem colunas sem HeadCtrl selecionado, selecione todos os HeadCtrls antes de prosseguir.", vbInformation
                    Exit Sub
                End If
                If lstSelectedColumns.Column(1, iInT) = dictQryListFields(iInT2) Then
                    sWidthColumns = sWidthColumns & lstSelectedColumns.Column(3, iInT) & ";"
                    bBoL = True
                End If
            Next iInT
            If Not bBoL Then
                sWidthColumns = sWidthColumns & "0;"
            End If
        Next iInT2

    cTrgtCtrl.ColumnWidths = sWidthColumns
    
    End If
    
    fTrgtForm.Repaint
End Sub

Private Sub btnPreview_Click()
    Dim sTrgtForm As String
    Dim sWidthColumns As String
    Dim iInT As Integer
    Dim iInT2 As Integer
    Dim bBoL As Boolean
    
    sTrgtForm = fTrgtForm.Name
If bAjustando = False Then
    btnPreview.Caption = "Fechar Preview"
    bAjustando = True
    Set cTrgtCtrl = fTrgtForm.Controls(txtSelectedList)
    sWidthColumns = ""
    
    For iInT2 = 1 To dictQryListFields.Count
        bBoL = False
        For iInT = 0 To lstSelectedColumns.ListCount - 1
            If lstSelectedColumns.Column(3, iInT) = "" Then
                MsgBox "Existem colunas sem HeadCtrl selecionado, selecione todos os HeadCtrls antes de prosseguir.", vbInformation
                Exit Sub
            End If
            If lstSelectedColumns.Column(1, iInT) = dictQryListFields(iInT2) Then
                sWidthColumns = sWidthColumns & lstSelectedColumns.Column(3, iInT) & ";"
                bBoL = True
            End If
        Next iInT
        If Not bBoL Then
            sWidthColumns = sWidthColumns & "0;"
        End If
    Next iInT2
    cTrgtCtrl.ColumnCount = dictQryListFields.Count
    cTrgtCtrl.ColumnWidths = sWidthColumns
    DoCmd.OpenForm sTrgtForm, acLayout
    Set fTrgtForm = Forms(sTrgtForm)
    
    Call Scr_FormAlwaysOnTop(Me, True)
Else
    btnPreview.Caption = "Preview"
    bAjustando = False
    DoCmd.OpenForm sTrgtForm, acDesign
    Set fTrgtForm = Forms(sTrgtForm)
End If
End Sub

Private Sub btnRemoveColumn_Click()
    Dim iInT As Integer
    Dim vA
    If lstSelectedColumns.ListIndex > -1 Then
        If dictSelectedFields.Exists(Int(lstSelectedColumns.Column(4, lstSelectedColumns.ListIndex))) Then _
        dictSelectedFields.Remove (Int(lstSelectedColumns.Column(4, lstSelectedColumns.ListIndex)))
        lstSelectedColumns.RemoveItem (lstSelectedColumns.ListIndex)
        
        lstFields.RowSource = ""
        For Each vA In dictQryListFields
            If Not dictSelectedFields.Exists(vA) Then _
            lstFields.AddItem dictQryListFields(vA), lstFields.ListCount
        Next vA
    End If
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
            If lstSelectedColumns.Column(3, iInT) = "" Then
                MsgBox "Existem colunas sem HeadCtrl selecionado, selecione todos os HeadCtrls antes de prosseguir.", vbInformation
                Exit Sub
            End If
            If lstSelectedColumns.Column(1, iInT) = dictQryListFields(iInT2) Then
                sWidthColumns = sWidthColumns & lstSelectedColumns.Column(3, iInT) & ";"
                bBoL = True
            End If
        Next iInT
        If Not bBoL Then
            sWidthColumns = sWidthColumns & "0;"
        End If
    Next iInT2
    cTrgtCtrl.ColumnCount = dictQryListFields.Count
    cTrgtCtrl.ColumnWidths = sWidthColumns
    bAjustando = False
    DoCmd.OpenForm fTrgtForm.Name, acDesign
    DoCmd.Close acForm, Me.Name, acSaveYes
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bAjustando Then
        Select Case KeyCode
            Case vbKeyAdd
                btnAdd_Click
                KeyCode = 0
                
            Case vbKeySubtract
                btnMin_Click
                KeyCode = 0
        End Select
    End If
End Sub

Private Sub Form_Load()
Dim fForm As Form

If Forms.Count > 2 Then
    MsgBox "Por favor deixe apenas o formulário que terá as listas manipuladas aberto antes de abrir este formulário.", vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End If
dictQryListFields.RemoveAll
dictSelectedFields.RemoveAll

bSelectedList = False
bSelectingHeads = False

For Each fForm In Forms
    If fForm.Name <> Me.Name Then Set fTrgtForm = fForm
Next fForm
End Sub

Private Sub Form_Timer()
On Error Resume Next
DoEvents

If bAjustando = False Then
    Set cTrgtCtrl = fTrgtForm.ActiveControl
    
    If IsNull(txtSelectedList) Then bSelectedList = False
    
    If Not Err.Number = 2474 Then
        
        DoCmd.SelectObject acForm, fTrgtForm.Name, False
        DoCmd.RunCommand acCmdSelectForm
        
        If sTrgtCtrl <> cTrgtCtrl.Name Then
            
            sTrgtCtrl = cTrgtCtrl.Name
            
            If cTrgtCtrl.ControlType = acListBox Then bSelectedList = False
            If lstSelectedColumns.ListIndex = -1 Then bSelectingHeads = False
            
            If bSelectedList = False And Not bSelectingHeads Then
                If cTrgtCtrl.ControlType = acListBox Then
                    dictQryListFields.RemoveAll
                    dictSelectedFields.RemoveAll
                    
                    lstSelectedColumns.RowSource = ""
                    bSelectedList = True
                    txtSelectedList = cTrgtCtrl.Name
                    lstFields.RowSource = ""
                    Set qDef = CurrentDb.QueryDefs(cTrgtCtrl.RowSource)
                    For Each fField In qDef.Fields
                        lstFields.AddItem fField.Name, lstFields.ListCount
                        dictQryListFields.Add lstFields.ListCount, fField.Name
                    Next fField
                    bSelectingHeads = True
                Else
                    txtSelectedList = "Selecione alguma List box"
                End If
                If lstFields.ListIndex > -1 Then lstFields.Selected(lstFields.ListIndex) = False
                Me.Move fTrgtForm.WindowLeft + cTrgtCtrl.Left + cTrgtCtrl.Width, cTrgtCtrl.Top + fTrgtForm.WindowTop
                
            End If
            iSelectedItem = lstSelectedColumns.ListIndex
            
            If iSelectedItem > -1 Then
                
                Call lstSelectedColumns_AfterUpdate
                
                If lstSelectedColumns.Column(3, lstSelectedColumns.ListIndex) <> "" Then
                    If iOldSelectedItem <> iSelectedItem Or iSelectedItem = 0 Then
                        If (lstSelectedColumns.ListIndex + 1) <= lstSelectedColumns.ListCount Then lstSelectedColumns.Selected(lstSelectedColumns.ListIndex + 1) = True
                    End If
                End If
            End If
        Else
            iSelectedItem = lstSelectedColumns.ListIndex
            
            If iSelectedItem > -1 Then
                
                Call lstSelectedColumns_AfterUpdate
                
                If lstSelectedColumns.Column(3, lstSelectedColumns.ListIndex) <> "" Then
                    If iOldSelectedItem <> iSelectedItem Or iSelectedItem = 0 Then
                        If (lstSelectedColumns.ListIndex + 1) <= lstSelectedColumns.ListCount Then lstSelectedColumns.Selected(lstSelectedColumns.ListIndex + 1) = True
                    End If
                End If
            End If
        End If
    End If
ElseIf bChanged Then
bChanged = False
 


End If


End Sub

Private Sub lstSelectedColumns_DblClick(Cancel As Integer)
Call btnRemoveColumn_Click
End Sub

Private Sub lstFields_DblClick(Cancel As Integer)
Call btnAddColumn_Click
End Sub

Private Sub lstSelectedColumns_AfterUpdate()
bSelectingHeads = True
Dim vDadosLst As Variant
Dim vA
Dim iInT As Integer, iInT2 As Integer, iConT As Integer
Dim sNewRowSource As String
If bAjustando = False Then
 iSelectedItem = lstSelectedColumns.ListIndex
    If cTrgtCtrl.ControlType = acCommandButton Then
        sHead = cTrgtCtrl.Caption
        If sHead <> sOldHead Then
        'lstSelectedColumns.Column(iSelectedItem, 2) Then
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
                
                
            ElseIf sHead <> lstSelectedColumns.Column(iSelectedItem, 2) <> "" Then

            End If
        
        End If

    End If
 Else
    txtWidth = lstSelectedColumns.Column(3, lstSelectedColumns.ListIndex)
 End If
    
End Sub


Private Sub txtWidth_Change()
    bChanged = True
End Sub


