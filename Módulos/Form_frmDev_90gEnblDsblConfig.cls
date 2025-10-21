VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90gEnblDsblConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim fTrgtForm As Form
Dim cTrgtCtrl As Control
Dim sTrgtCtrl As String
Dim bSelectedList As Boolean
Dim iSelectedItem As Integer
Dim sSQL As String
Dim rsQry As Recordset
Dim dictFormModes As New Dictionary
Dim sID As String


Private Sub btnCancelar_Click()
Dim iInT As Integer
Do While lstCtrls.ListCount > 0
    lstCtrls.Selected(0) = True
    Call RemoveRegistro
Loop
End Sub

Private Sub btnClose_Click()
DoCmd.Close

End Sub

Private Sub chkEnableCtrl_Click()
Call AtualizarValores(lstCtrls.ListIndex)
End Sub

Private Sub chkLockedCtrl_Click()
Call AtualizarValores(lstCtrls.ListIndex)
End Sub

Private Sub chkVisibleCtrl_Click()
Call AtualizarValores(lstCtrls.ListIndex)
End Sub

Private Sub cmbForm_AfterUpdate()
 Call FiltraQryEnblDsbl
End Sub



Private Sub cmbFormModes_AfterUpdate()
Dim vA, vB, vC
Dim iInT As Integer
Dim bBoL As Boolean

bBoL = False

For iInT = 0 To cmbFormModes.ListCount - 1
    If cmbFormModes.Column(0, iInT) = cmbFormModes.Value Then bBoL = True
Next iInT
If Not bBoL Then
    cmbFormModes.AddItem cmbFormModes.Value
End If

Call FiltraQryEnblDsbl
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If bAjustando Then
'        Select Case KeyCode
'            Case vbKeyAdd
'                btnAdd_Click
'                KeyCode = 0
'
'            Case vbKeySubtract
'                btnMin_Click
'                KeyCode = 0
'        End Select
'    End If
End Sub

Private Sub Form_Load()
Dim fForM As Form
Dim lbl As Label
Dim objForm As Object
Dim vA
If Forms.Count > 2 Then
    MsgBox "Por favor deixe apenas o formulário que terá as listas manipuladas aberto antes de abrir este formulário.", vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End If

bSelectedList = False

For Each fForM In Forms
    If fForM.Name <> Me.Name Then Set fTrgtForm = fForM
Next fForM

If Not fTrgtForm Is Nothing Then
    cmbForm.Value = fTrgtForm.Name
Else

For Each objForm In CurrentProject.AllForms
    cmbForm.AddItem objForm.Name
Next objForm
cmbForm.Value = cmbForm.ItemData(0)
    
End If
'qry_01(03)bSysEnblDisblParams
Set rsQry = CurrentDb.OpenRecordset("qry_01(03)bSysEnblDisblParams", dbOpenDynaset, dbReadOnly)
rsQry.MoveFirst
dictFormModes.RemoveAll

Do While Not rsQry.EOF
    If dictFormModes.Exists(rsQry("sSysFormMode").Value) Then dictFormModes.Remove (rsQry("sSysFormMode").Value)
    dictFormModes.Add rsQry("sSysFormMode").Value, rsQry("sSysFormMode")
    rsQry.MoveNext
Loop
rsQry.Close
Set rsQry = Nothing
cmbFormModes.RowSource = ""
For Each vA In dictFormModes
    cmbFormModes.AddItem vA
Next vA

cmbFormModes.Value = cmbFormModes.ItemData(0)

Call FiltraQryEnblDsbl

lstCtrls.SetFocus
End Sub

Private Sub FiltraQryEnblDsbl()
Dim sFilterMode As String
Dim sFilterForm As String
Dim sWhere As String
Dim qDef As QueryDef

Set qDef = CurrentDb.QueryDefs("qry_01(03)bSysEnblDisblParams")
sSQL = qDef.sql

sSQL = Split(sSQL, "ORDER BY")(0)
sFilterForm = "([sSysForM] = " & """" & cmbForm.Value & """" & ")"
sFilterMode = "([sSysFormMode] = " & """" & cmbFormModes.Value & """" & ")"

sWhere = "WHERE " & sFilterForm & " AND " & sFilterMode

sSQL = sSQL & vbCrLf & sWhere

lstCtrls.RowSource = sSQL

End Sub


Private Sub Form_Timer()
Dim sIDValue As String
Dim vA, vB, vC
Dim fForM As Form
Dim qDef As QueryDef
Dim sSQL As String
Dim sWhere As String

Set qDef = CurrentDb.QueryDefs("qry_01(03)bSysEnblDisblParams")
sSQL = qDef.sql

On Error Resume Next
DoEvents

If Screen.ActiveForm.Name <> fTrgtForm.Name Then
    Me.TimerInterval = 1000
    GoTo SkipTimer
Else
    
    If Not fTrgtForm Is Nothing Then
    
        Set cTrgtCtrl = fTrgtForm.ActiveControl
        
        DoCmd.SelectObject acForm, fTrgtForm.Name, False
        DoCmd.RunCommand acCmdSelectForm
        
        If Not cTrgtCtrl Is Nothing Then
            If sTrgtCtrl <> cTrgtCtrl.Name Then
            
                sTrgtCtrl = cTrgtCtrl.Name
                
                sSQL = qDef.sql
                sSQL = Split(sSQL, "ORDER BY")(0)
                vA = "([sSysForM] = " & """" & cmbForm.Value & """" & ")"
                vB = "([sSysFormMode] = " & """" & cmbFormModes.Value & """" & ")"
                vC = "([sTweakbleCtrl] = " & """" & sTrgtCtrl & """" & ")"
                sWhere = "WHERE " & vA & " AND " & vB & " AND " & vC
                sSQL = sSQL & vbCrLf & sWhere & " ORDER BY sTweakbleCtrl"
                
                Set rsQry = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)
                
                If rsQry.RecordCount = 0 Then
                    Call AdicionaRegistro
                    lstCtrls.Requery
                End If
            End If
        End If
    Else
        For Each fForM In Forms
            If fForM.Name <> Me.Name Then Set fTrgtForm = fForM
        Next fForM
    End If
End If
SkipTimer:

End Sub

Private Sub lstCtrls_AfterUpdate()
Dim iLstIndex As Integer

iLstIndex = lstCtrls.ListIndex

chkEnableCtrl.Value = lstCtrls.Column(6, iLstIndex)
chkVisibleCtrl.Value = lstCtrls.Column(7, iLstIndex)
chkLockedCtrl.Value = lstCtrls.Column(8, iLstIndex)
txtTipTextCtrl = lstCtrls.Column(9, iLstIndex)


End Sub
Sub RemoveRegistro()
Dim sIDValue As String
Dim qDef As QueryDef
Dim sSQL As String
Dim sWhere As String
Dim iLstIndex As Integer
Dim vA, vB, vC
Set qDef = CurrentDb.QueryDefs("qry_01(03)bSysEnblDisblParams")
sSQL = qDef.sql
iLstIndex = lstCtrls.ListIndex
sSQL = Split(sSQL, "ORDER BY")(0)
sID = qDef.Fields(0).Name
sIDValue = lstCtrls.Column(0, iLstIndex)

vA = "([" & sID & "] = " & sIDValue & ")"

sWhere = "WHERE " & vA
sSQL = sSQL & vbCrLf & sWhere

Set rsQry = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)

rsQry.Delete
'rsQry.Update

rsQry.Close
Set rsQry = Nothing
lstCtrls.Requery

End Sub
Sub AdicionaRegistro()
Set rsQry = CurrentDb.OpenRecordset("qry_01(03)bSysEnblDisblParams", dbOpenDynaset)
rsQry.AddNew
rsQry("sSysForM") = cmbForm.Value
rsQry("sSysFormMode") = cmbFormModes.Value
rsQry("sTweakbleCtrl") = sTrgtCtrl

rsQry("bEnable") = chkEnabled.Value
rsQry("bVisible") = chkVisible.Value
rsQry("bLockCombo") = chkLocked.Value

rsQry.Update

rsQry.Close
Set rsQry = Nothing
End Sub
Sub AtualizarValores(iLstIndex As Integer)
Dim sIDValue As String
Dim qDef As QueryDef
Dim sSQL As String
Dim sWhere As String
Dim vA, vB, vC
Set qDef = CurrentDb.QueryDefs("qry_01(03)bSysEnblDisblParams")
sSQL = qDef.sql

sSQL = Split(sSQL, "ORDER BY")(0)
sID = qDef.Fields(0).Name
sIDValue = lstCtrls.Column(0, iLstIndex)

vA = "([" & sID & "] = " & sIDValue & ")"

sWhere = "WHERE " & vA
sSQL = sSQL & vbCrLf & sWhere

Set rsQry = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)

rsQry.Edit
rsQry("bEnable") = chkEnableCtrl.Value
rsQry("bVisible") = chkVisibleCtrl.Value
rsQry("bLockCombo") = chkLockedCtrl.Value
rsQry("sAltTipText") = txtTipTextCtrl
rsQry.Update

rsQry.Close
Set rsQry = Nothing

lstCtrls.Requery

End Sub


Private Sub lstCtrls_DblClick(Cancel As Integer)
Call RemoveRegistro
End Sub

Private Sub txtTipTextCtrl_LostFocus()
Call AtualizarValores(lstCtrls.ListIndex)
End Sub
