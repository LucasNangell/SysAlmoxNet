VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90cDevListBoxColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSolicitantes_Click()
    Dim sOpenArgs As String, sOpenArgs1 As String, sOpenArgs2 As String
    Dim iCoD As Integer
    Dim sStR1 As String, sStR2 As String
    
    sStR1 = "frmSolicitantes1CadSearch"   'formulário que deve ser aberto
    sOpenArgs1 = "lstSolicitantes"        'campo no novo form que tem o valor a ser selecionado
    sOpenArgs2 = "txtSrchSolicit"         'controle que deve receber o foco após abertura do form
    iCoD = 0   'Item a ser selecionado no Form que será aberto
    sOpenArgs = sOpenArgs1 & "|" & iCoD & "|" & sOpenArgs2

'Stop
    
    DoCmd.Minimize
    DoCmd.OpenForm sStR1, , , , , , sOpenArgs


End Sub

Private Sub btnSair_Click()
    DoCmd.Close acForm, Me.Name, acSavePrompt

End Sub

Private Sub btnMinimizar_Click()
    
    'Reduzir a janela do Access e posicionar pro centro da tela, pra ficar por trás do Form
    DoCmd.RunCommand acCmdAppRestore
    Call AccessMoveSize(500, 300, 500, 300)   'Left, Top, Width, Heigth

End Sub




Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name, acSavePrompt
    
End Sub

Private Sub Form_Load()
    Dim vA, vB
    Dim vKeyCtrl As Variant
    Dim iCount As Integer
    Dim sFontName As String
    Dim sCtrL As String
    Dim sListboxColor As String
    
'Stop
    For Each vKeyCtrl In Me.Controls
        sCtrL = vKeyCtrl.Name
        
        vA = vKeyCtrl.ControlType
        If vA = acListBox Then
'Stop
            vKeyCtrl.Selected(2) = True
            
            vB = InStr(sCtrL, "lstFont")
            If vB > 0 Then
                iCount = iCount + 1
                sFontName = Me.Controls("lstFont" & Format(iCount, "00")).FontName
                sListboxColor = Me.Controls("lstFont" & Format(iCount, "00")).BackColor
'Stop
                
                Me.Controls("txtFonte" & Format(iCount, "00")) = sFontName & "          BackColor: " & sListboxColor
            
            End If
            
        End If
        
        
    Next vKeyCtrl
    
End Sub
