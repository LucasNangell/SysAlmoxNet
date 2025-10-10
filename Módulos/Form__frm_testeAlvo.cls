VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__frm_testeAlvo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSair_Click()

'Stop
    Call Scr_ShowAccess
    DoCmd.Close
    
End Sub

Private Sub tglDockingTrggFrm_Click()
    Set clObjFormOpenParams = New cls_09cParamsToOpenForms
    
    With clObjFormOpenParams
        Set .cTglBtnDocking = Me.tglDockingTrggFrm
        Set .fTrgtForm = Me
        .sTrgtForm = Me.Name
        .bFrmIsDocked = .cTglBtnDocking.Value
        
    End With
    
    Call FormLoad02b_UpdateFormPositionDict(clObjFormOpenParams)
End Sub

Private Sub tglFixedOnTop_Click()
    Dim cFixOnTop As Control
    
    Set cFixOnTop = Me.tglFixedOnTop

    Call Scr_FormAlwaysOnTop(Me, cFixOnTop.Value)
    
End Sub

