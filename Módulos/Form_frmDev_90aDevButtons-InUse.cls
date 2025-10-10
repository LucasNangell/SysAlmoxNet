VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90aDevButtons-InUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCor_Click()
    
    Dim CtlColorLNG As Long
    Dim CtlGradientINT As Integer
    Dim CtlBdColorLNG As Long
    
    Dim LongRGB As Long
    Dim StringRGB As String
    Dim StringHEX As String
    
    Dim CorPadrao As Long
    Dim PickedColor As Long
    
    Dim CoRhexFunc As MyColors
    
    Dim CorAccessGotLNG As Long
    Dim CorStrGotSTR As String
    Dim iPartColor As Integer

    Dim vA, vB, vC

'Stop

    'sim = 6, não = 7, cancel = 2
    iPartColor = MsgBox("Sim - para Back color" & vbCr & "Não - para Fore color" & vbCr & "Cancel - para Border color", vbYesNoCancel)
'Stop
    'If iPartColor = 2 Then Exit Sub
    CtlColorLNG = IIf(iPartColor = 6, btnMudar.BackColor, IIf(iPartColor = 7, btnMudar.ForeColor, btnMudar.BorderColor))
    'CtlColorLNG = IIf(iPartColor = 6, btnMudar.BackColor, btnMudar.BorderColor)
    
    
    'CtlColorLNG = btnMudar.BackColor
    'CtlColorLNG = btnMudar.forecolor
    'CtlBdColorLNG = btnMudar.BorderColor
    'CtlGradientINT = btnMudar.Gradient
'Stop
    PickedColor = ColorPicker(CtlColorLNG)
'Stop
    'If PickedColor = 0 Then Exit Sub
    
    Select Case iPartColor
    
        Case 6  'Back color
            btnMudar.BackColor = PickedColor
            btnMudar.Gradient = igbBtNGdntColor
            
        Case 7  'Fore color
            btnMudar.ForeColor = PickedColor
        
        Case 2  'Border color
            btnMudar.BorderColor = PickedColor 'igbBtNclearForeColor
    
    End Select
    
    
'----
    'obtém os valores RGB e HEX a partir do código Long interno do Access
    CoRes = DecToMyColorsRGB_HEX(PickedColor)
    'CoRes = DecToMyColorsRGB_HEX(CtlColorLNG)
    
    LongRGB = CoRes.CorAccessGBlng   'longRGB
    StringRGB = CoRes.CorRGBgbStr      'string RGB
    StringHEX = CoRes.CorHEXgbStr      'string HEX
'----
    
'Stop
    
    MsgBox "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
'    Debug.Print "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
    
'Stop
    
'    btnNewColor.BackColor = HexToLongRGB(sgbBtnGREyBackColor)
'    btnNewColor.BorderColor = HexToLongRGB(sgbBtnGREyBdColor)
'    btnNewColor.Gradient = igbBtNGdntColor
'    btnNewColor.ForeColor = igbBtNclearForeColor
'    btnNewColor.Enabled = False
'
'    btnNewColor.BackColor = HexToLongRGB(sgbBtnBLUeBackColor)
'    btnNewColor.BorderColor = HexToLongRGB(sgbBtnBLUeBdColor)
'    btnNewColor.Gradient = igbBtNGdntColor
'    btnNewColor.ForeColor = igbBtNclearForeColor
'
'    btnNewColor.BackColor = HexToLongRGB(sgbBtnREdBackColor)
'    btnNewColor.BorderColor = HexToLongRGB(sgbBtnREdBdColor)
'    btnNewColor.Gradient = igbBtNGdntColor
'    btnNewColor.ForeColor = igbBtNclearForeColor
'
'    btnNewColor.BackColor = igbBtnYELLowBackColor
'    btnNewColor.BorderColor = igbBtnYELLowBdColor
'    btnNewColor.Gradient = igbBtNGdntColor
'    btnNewColor.ForeColor = igbBtNclearForeColor
'
'    btnNewColor.BackColor = HexToLongRGB(sgbBtnGREEnBackColor)
'    btnNewColor.BorderColor = HexToLongRGB(sgbBtnGREEnBdColor)
'    btnNewColor.Gradient = igbBtNGdntColor
'    btnNewColor.ForeColor = igbBtNclearForeColor
'
'
'
'
'
'    btnMudar.BackColor = PickedColor
'    btnMudar.Gradient = CtlGradientINT
'    'btnMudar.BorderColor
'    btnMudar.Enabled = False
    
'    btnMudar.ForeColor
'    btnMudar.ForeShade
'    btnMudar.ForeThemeColorIndex
'    btnMudar.ForeTint
'    btnMudar.Gradient
'    btnMudar.HoverColor
'    btnMudar.HoverForeColor
'    btnMudar.HoverForeShade
'    btnMudar.HoverForeThemeColorIndex
'    btnMudar.HoverForeTint
'    btnMudar.HoverForeShade
'    btnMudar.HoverForeThemeColorIndex
'    btnMudar.HoverTint
'    btnMudar.ObjectPalette
'    btnMudar.PressedColor
'    btnMudar.PressedForeColor
'
'Stop


End Sub


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
