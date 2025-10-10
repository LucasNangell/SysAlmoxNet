VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90bDevChangeCtlColors - apagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCheckColor_Click()
    
    Dim PickedColor As Long
    Dim lHEX As Long
    Dim CtlColorLNG As Long
    Dim LongRGB As Long
    Dim StringRGB As String
    Dim StringHEX As String
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    
    Dim CtlGradientINT As Integer
    Dim CtlBdColorLNG As Long
    
    Dim CorPadrao As Long
    
    Dim CoRhexFunc As MyColors
    
    Dim CorAccessGotLNG As Long
    Dim CorStrGotSTR As String

    Dim a, B, c
    Dim Tmp
    Dim iPartColor


'Stop

    iPartColor = MsgBox("Sim - para Back color" & vbCr & "Não - para Border color", vbYesNoCancel)
'Stop
    If iPartColor = 2 Then Exit Sub
    CtlColorLNG = IIf(iPartColor = 6, btnNewColor.BackColor, btnNewColor.BorderColor)
'Stop
    'a= HSLtoRGB(
    'HSLtoRGB = rgb(red, green, blue)
'    a = btnNewColor.Hue
'    a = btnNewColor.Sat
'    a = btnNewColor.Lum
    
    'CtlColorLNG = btnNewColor.BackColor
    'CtlColorLNG = btnNewColor.forecolor
    'CtlBdColorLNG = btnNewColor.BorderColor
    'CtlGradientINT = btnNewColor.Gradient
    
    PickedColor = ColorPicker(CtlColorLNG)
'Stop
    
    'obtém os valores RGB e HEX a partir do código Long interno do Access
    CoRes = DecToMyColorsRGB_HEX(PickedColor)
    'CoRes = DecToMyColorsRGB_HEX(CtlColorLNG)
    
    LongRGB = CoRes.CorAccessGBlng   'longRGB
    StringRGB = CoRes.CorRGBgbStr      'string RGB
    StringHEX = CoRes.CorHEXgbStr      'string HEX
    
'Stop
    MsgBox "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
'    Debug.Print "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
    
'Stop
'    GetHSL = MyRGBtoHSL(GetRGB.Red, GetRGB.Green, GetRGB.Blue) ' funcionou
'
'    'Avaliação do HSL
'
'    a = GetHSL.lHue
'    a = GetHSL.lLum
'    If GetHSL.lHue < 11 Or (GetHSL.lHue > 129 And GetHSL.lHue < 240) Then a = "Botão Escuro"
'    If GetHSL.lHue > 21 And GetHSL.lHue < 130 Then a = "Botão Claro"
    

'Stop
    
Exit Sub
    
    
'
'
'    GetHSL = GetHSLfromRGB2(GetRGB.Red, GetRGB.Green, GetRGB.Blue) ' funcionou
'    GetHSL = HSLtoRGB240(GetRGB.Red, GetRGB.Green, GetRGB.Blue)
'
'    GetHSL = RgbToHLSstack(GetRGB.Red, GetRGB.Green, GetRGB.Blue)
'    GetHSL = Rgb2Hsl(GetRGB.Red, GetRGB.Green, GetRGB.Blue)
'    GetHSL = Rgb2Hsl_2(GetRGB.Red, GetRGB.Green, GetRGB.Blue)
'
'
'
'
'    HSL3in1 = RGBToHSL03(PickedColor)
'    HSL3in1 = RGBToHSL02(PickedColor)
'
'    HSL3in1 = RGBToHSL01(PickedColor)
'
'
'    GetHSL = GetHSLfromRGB2(GetRGB.Red, GetRGB.Green, GetRGB.Blue)
'
'
'    lRed = GetRGB.Red
'    lGreen = GetRGB.Green
'    lBlue = GetRGB.Blue
'
'    'TmP = GetHSLFromRGB(lRed, lGreen, lBlue)
'
'    'obtém o valor RGB Long a partir do HEX, necessário pra atribuir cores para os controles no VBA
'    LongRGB = HexToLongRGB(CoRes.CorHEXgbStr)
'
'
'Stop
'
'    'GetRGB = HEXtoRGB(CoRes.CorHEXgbStr)
'
'
'
'
'
'
'    'GetRGB = DECtoRGB(PickedColor)
'
'    'CoRes = DECtoMyColorsRGB_HEX(PickedColor)
'
'Stop
'    'CoRes = LongDecToRGB_HEX_HSL(PickedColor)
'    'CoRes = GetHSLfromRGB(PickedColor)
'Stop
'
'
'    MsgBox "Cor selecionada:" & vbCr & vbCr & _
'        " - em valor Decimal: " & CoRes.CorAccessGBlng & vbCr & _
'        " - em valor RGB:     " & CoRes.CorRGBgbStr & vbCr & _
'        " - em valor HEX:     " & CoRes.CorHEXgbStr & vbCr & _
'        " - em valor HLS:     " & CoRes.CorHSLgbStr
'
'
'    If PickedColor = 0 Then Exit Sub
'    btnNewColor.BackColor = PickedColor
'    btnNewColor.ForeColor = PickedColor - 2450350
'    btnNewColor.BorderColor = sgbBtnGREyBdColor
'    btnNewColor.Gradient = igbBtNGdntColor
'
'    btnNewColor.PressedColor = 21
'    btnNewColor.PressedForeColor = 21
'    btnNewColor.HoverForeColor = 21
'    btnNewColor.HoverColor = 21
'    btnNewColor.HoverForeTint = 21
'
'
'Stop
'
''    btnNewColor.BackColor = BtNgreyBackColor
''    btnNewColor.BorderColor = BtNgreyBdColor
''    btnNewColor.Gradient = igbBtNGdntColor
''    btnNewColor.ForeColor = igbBtNclearForeColor
''    btnNewColor.Enabled = False
''
''    btnNewColor.BackColor = BtNblueBackColor
''    btnNewColor.BorderColor = BtNblueBdColor
''    btnNewColor.Gradient = igbBtNGdntColor
''    btnNewColor.ForeColor = igbBtNclearForeColor
''
''    btnNewColor.BackColor = BtNredBackColor
''    btnNewColor.BorderColor = BtNredBdColor
''    btnNewColor.Gradient = igbBtNGdntColor
''    btnNewColor.ForeColor = igbBtNclearForeColor
''
''    btnNewColor.BackColor = BtNyellowBackColor
''    btnNewColor.BorderColor = BtNyellowBdColor
''    btnNewColor.Gradient = igbBtNGdntColor
''    btnNewColor.ForeColor = igbBtNclearForeColor
''
''    btnNewColor.BackColor = BtNgreenBackColor
''    btnNewColor.BorderColor = BtNgreenBdColor
''    btnNewColor.Gradient = igbBtNGdntColor
''    btnNewColor.ForeColor = igbBtNclearForeColor
'
'
'
'
'
'    btnNewColor.BackColor = PickedColor
'    btnNewColor.Gradient = CtlGradientINT
'    'btnNewColor.BorderColor
'    btnNewColor.Enabled = False



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
    DoCmd.Close acForm, Me.Name, acSavePrompt
    'DoCmd.Quit

End Sub

