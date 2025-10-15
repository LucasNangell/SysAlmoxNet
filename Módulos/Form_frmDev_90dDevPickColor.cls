VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDev_90dDevPickColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'--------------------------------------------------------------------------------------
'------------------------------------------------
'Declaração de variáveis pro Form [ ColorPicker ]
    Private mArrastando As Boolean
    Private mLastX As Single
    Private mLastY As Single
    Private lngOriginalColor As Long
    'Private bManterCor As Boolean

    Dim fForM As Form
    Dim cTrgtCtrl As Control
    Dim cVertBarLevels As Control


''--------------------------------------------------------------------------------------
''------------------------------------------------
''Declaração de variáveis pras funções do Form [ ColorPicker ]

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, _
                                                           ByVal x As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
#End If


Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

' Constante para o cursor de cruz
Private Const IDC_CROSS = 32515
Private Const IDC_ARROW = 32512     ' Cursor normal (seta)


Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type RGB
    R As Integer
    G As Integer
    B As Integer

End Type

'Dim cVertBarLevels As Control
''------------------------------------------------
''--------------------------------------------------------------------------------------


Private Sub btnSalvar_Click()
    'bManterCor = True
    
    If Not cTrgtCtrl Is Nothing Then cTrgtCtrl.SetFocus
    DoCmd.Close
    
End Sub


Private Sub btnCancelar_Click()
    
    Dim vA
    
    
    If Not cTrgtCtrl Is Nothing Then
        cTrgtCtrl.BackColor = lngOriginalColor
        cTrgtCtrl.SetFocus
    
    End If
    
    DoCmd.Close

End Sub

Private Sub Form_Load()
    
    Dim vA, vB
    Dim tRGB As RGB
    Dim bBoL As Boolean
    
    vA = CurrentProject.AllForms(TempVars("sForM")).IsLoaded
    If Not vA Then Exit Sub
    
    Me.btnSalvar.SetFocus
    Set fForM = Forms(TempVars("sForM"))
    Set cTrgtCtrl = fForM.Controls(TempVars("sTrgtCtrL"))
    
    'bManterCor = False
    lngOriginalColor = cTrgtCtrl.BackColor
    boxSeletorHB.Left = imgHB.Left - ((boxSeletorHB.Width - imgHB.Width) / 2)
    boxSeletorHB.Top = imgHB.Top + (imgHB.Height / 2)
    'Set cCtrL = Forms(TempVars("sForm")).Controls(TempVars("sTrgtCtrl"))
    'chkViewRealTime.Value = True
    
    Me.txtColorCheck.Value = cTrgtCtrl.Name
    Me.Move fForM.WindowLeft + cTrgtCtrl.Left + cTrgtCtrl.Width + 150, fForM.WindowTop + cTrgtCtrl.Top
    
    
    Call AtualizaCorDoPonto(Me, cTrgtCtrl, lngOriginalColor)

End Sub

Private Sub boxSeletorHB_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button = 1 Then                      ' Botão esquerdo
        mArrastando = True
        mLastX = x
        mLastY = Y

    End If
End Sub

Private Sub boxSeletorHB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Dim dx As Single, dy As Single
        Dim dTam As Double
        Dim iPercTop As Integer
        Dim cVertBar As Control
        Dim tRGB As RGB
        Dim minL As Single, minT As Single, maxL As Single, maxT As Single
        
    If mArrastando Then
        dx = x - mLastX                      ' delta desde o último evento
        dy = Y - mLastY

        With Me.boxSeletorHB
            '.Left = .Left + dx               ' move o controle
            .Top = .Top + dy

            ' --- Limitar dentro da área visível do formulário ---
            
            minT = Me.imgHB.Top
            maxT = LimitLineDown.Top
            dTam = maxT - minT
            
            If .Top < minT Then
                .Top = minT
                mLastX = x
                mLastY = Y
            End If
            If .Top > maxT Then
                .Top = maxT
                mLastX = x
                mLastY = Y
            End If
            
            
            iPercTop = Int(100 * ((boxSeletorHB.Top - imgHB.Top) / dTam) / 2)
            If iPercTop > 50 Then iPercTop = 50
            If iPercTop < 1 Then iPercTop = 1
            
            If ControlExists("bx" & iPercTop, Me) Then
                Set cVertBar = Me.Controls("bx" & iPercTop)
                boxPreview.BackColor = cVertBar.BackColor
                txtHex = ConverterCorVBAparaHex(boxPreview.BackColor)
                tRGB = HexToRGB(txtHex)
                txtR = tRGB.R
                txtG = tRGB.G
                txtB = tRGB.B
                'If chkViewRealTime.Value = True Then
                    If Not cTrgtCtrl Is Nothing Then
                    cTrgtCtrl.BackColor = cVertBar.BackColor
                    Forms(cTrgtCtrl.Parent.Name).Repaint
                    
                    End If
            
                'End If
            
            End If
        
        End With
    
    End If
    
End Sub

Private Sub boxSeletorHB_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        mArrastando = False
        mLastX = x
        mLastY = Y
    End If
End Sub


Private Sub imgSV_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim vA

    If Button = 1 Then
            
            SetCursor (LoadCursor(0, IDC_CROSS))
            
            mArrastando = True
            'If Y > imgSV.Top Then Debug.Print Y
            Call AtualizaCorDoPonto(Me, cTrgtCtrl)
      
    End If
    Dim pt As POINTAPI

    
    vA = GetCursorPos(pt)
    vA = Me.WindowLeft
    vA = Me.WindowTop
    
End Sub

Private Sub imgSV_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim dx As Single, dy As Single
    
    dx = x - mLastX                      ' delta desde o último evento
    dy = Y - mLastY

    If mArrastando Then
        If x >= 0 And x <= imgSV.Width And Y >= 0 And Y <= imgSV.Height Then Call AtualizaCorDoPonto(Me, cTrgtCtrl)
    Else
        'volta cursor ao normal
        SetCursor (LoadCursor(0, IDC_ARROW))
    End If
End Sub

Private Sub imgSV_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        mArrastando = False
        mLastX = x
        mLastY = Y
    End If
End Sub

Private Sub imgSV_Click()
    ' Também pega a cor num clique simples, se preferir
    Call AtualizaCorDoPonto(Me, cTrgtCtrl)
End Sub


'--------------------------------------------------------------------------------------
'------------------------------------------------
'Funções e Rotinas usadas pelo Form [ ColorPicker ]
'------------------------------------------------
'--------------------------------------------------------------------------------------

' Retorna a cor (OLE_COLOR) do pixel na posição de tela (screenX, screenY)
Private Function PegaCorPixel(ByVal screenX As Long, ByVal screenY As Long) As Long
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If
    Dim clr As Long

    hdc = GetDC(0)                     ' DC da tela inteira
    clr = GetPixel(hdc, screenX, screenY)
    Call ReleaseDC(0, hdc)

    PegaCorPixel = clr                 ' COLORREF (bbggrr) compatível com OLE_COLOR
End Function

' Atalho: pega a cor sob o cursor **agora**, sem precisar calcular nada
Private Function PegaCorDoCursor() As Long
    Dim pt As POINTAPI
    If GetCursorPos(pt) <> 0 Then
        PegaCorDoCursor = PegaCorPixel(pt.x, pt.Y)
    Else
        PegaCorDoCursor = -1
    End If
End Function

' Utilitário: devolve texto em HEX (ex.: #12A3F0)
Private Function ColorToHex(ByVal c As Long) As String
    ColorToHex = "#" & Right$("000000" & hex(c And &HFFFFFF), 6)
End Function

' Converte "#RRGGBB" para OLE_COLOR (COLORREF = 0x00BBGGRR)
Private Function HexToColor(ByVal hex As String) As Long
    Dim tRGB As RGB
    tRGB = HexToRGB(hex)

    HexToColor = (tRGB.B * &H10000) Or (tRGB.G * &H100) Or tRGB.R
End Function

Private Function MixColor(ByVal hexBase As String, ByVal Perct As Double) As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim pR As Double, pG As Double, pB As Double
    Dim dQtStep As Double
    Dim tRGB As RGB
    tRGB = HexToRGB(hexBase)

    If Perct > 100 Then Perct = 100
    If Perct < 50 Then dQtStep = ((Perct * 2.08) / 100) * 25
    If Perct > 50 Then dQtStep = (((Perct - 50) * 2.08) / 100) * 25

    If Perct < 50 Then
        pR = (255 - tRGB.R) / 25
        pG = (255 - tRGB.G) / 25
        pB = (255 - tRGB.B) / 25

        newR = 255 - (pR * dQtStep) ' / 100)
        newG = 255 - (pG * dQtStep) ' / 100)
        newB = 255 - (pB * dQtStep) ' / 100)

    ElseIf Perct = 50 Then
        newR = tRGB.R
        newG = tRGB.G
        newB = tRGB.B

    ElseIf Perct > 50 Then
        pR = tRGB.R / 25
        pG = tRGB.G / 25
        pB = tRGB.B / 25

        newR = tRGB.R - ((pR * dQtStep))
        newG = tRGB.G - ((pG * dQtStep))
        newB = tRGB.B - ((pB * dQtStep))

    End If

    MixColor = RGB(newR, newG, newB)

End Function
Private Function HexToRGB(hexValue As String) As RGB
    Dim s As String, R As Long, G As Long, B As Long
    s = Replace(Replace(Replace(Trim$(hexValue), "#", ""), "0x", ""), "&H", "")
    's = Replace(Replace(Replace(Trim$(hex), "#", ""), "0x", ""), "&H", "")
    If Len(s) <> 6 Then Err.Raise 5, , "Hex inválido (use #RRGGBB)"
    HexToRGB.R = CLng("&H" & Mid$(s, 1, 2))
    HexToRGB.G = CLng("&H" & Mid$(s, 3, 2))
    HexToRGB.B = CLng("&H" & Mid$(s, 5, 2))
End Function
Function ConverterCorVBAparaHex(ByVal corVBA As Long) As String

    Dim R As Long
    Dim G As Long
    Dim B As Long

    R = corVBA Mod 256
    G = (corVBA \ 256) Mod 256
    B = (corVBA \ 65536) Mod 256

    ' Concatena os valores hexadecimais, garantindo que cada um tenha 2 dígitos.
    ConverterCorVBAparaHex = "#" & Right("00" & hex(R), 2) & Right("00" & hex(G), 2) & Right("00" & hex(B), 2)
End Function

Private Function HexToOleColor(ByVal hex As String) As Long
    Dim s As String, R As Long, G As Long, B As Long

    s = UCase$(Trim$(hex))
    s = Replace(s, "#", "")
    s = Replace(s, "&H", "")
    s = Replace(s, "0X", "")

    ' suporta #RGB (shorthand CSS)
    If Len(s) = 3 Then
        s = Mid$(s, 1, 1) & Mid$(s, 1, 1) & _
            Mid$(s, 2, 1) & Mid$(s, 2, 1) & _
            Mid$(s, 3, 1) & Mid$(s, 3, 1)
    End If

    If Len(s) <> 6 Then Err.Raise 5, , "Hex inválido: use #RRGGBB ou #RGB"

    R = CLng("&H" & Mid$(s, 1, 2))
    G = CLng("&H" & Mid$(s, 3, 2))
    B = CLng("&H" & Mid$(s, 5, 2))

    HexToOleColor = RGB(R, G, B) ' <- formato aceito por BackColor
End Function


Private Sub AtualizaCorDoPonto(fForM As Form, cTrgtCtrl As Control, Optional lngCor As Long)
    Dim iBox As Integer
    Dim iPrct As Double
    Dim iPercTop As Integer
    Dim iConT As Integer
    Dim cCtrlH As Control
    Dim sColor As String
    Dim sColorVa As String
    Dim lngColorVa As Long
    Dim R, G, B
    Dim dTam As Double
    Dim tRGB As RGB
    
    iConT = 0
    iPrct = iConT * 2
    
    If lngCor = 0 Then lngCor = PegaCorDoCursor()
    sColor = ConverterCorVBAparaHex(lngCor)

    If lngCor <> -1 Then
        'Me!boxPreview.BackColor = lngCor
        dTam = fForM.LimitLineDown.Top - fForM.imgHB.Top
        
        iPercTop = Int(100 * ((fForM.boxSeletorHB.Top - fForM.imgHB.Top) / dTam) / 2)
        If iPercTop > 50 Then iPercTop = 50
        If iPercTop < 1 Then iPercTop = 1
        fForM.boxPreview.BackColor = lngCor
        
        If ControlExists("bx" & iPercTop, fForM) Then
            Set cCtrlH = fForM.Controls("bx" & iPercTop)
            cCtrlH.BackColor = lngCor
            'fForM.boxPreview.BackColor = cCtrlH.BackColor
            'If chkViewRealTime.Value = True Then
                If Not cTrgtCtrl Is Nothing Then
                    cTrgtCtrl.BackColor = cCtrlH.BackColor
                    Forms(cTrgtCtrl.Parent.Name).Repaint
                    
                End If
            
            'End If
            fForM.txtHex.Value = ConverterCorVBAparaHex(fForM.boxPreview.BackColor)
            fForM.txtOLEcolor.Value = lngCor
'Stop
        End If
            
        If Not fForM!txtHex Is Nothing Then
            fForM!txtHex.Value = sColor
            tRGB = HexToRGB(sColor)
            fForM.txtR = tRGB.R
            fForM.txtG = tRGB.G
            fForM.txtB = tRGB.B
            iConT = 0
            
            For iBox = 1 To 50
                Set cVertBarLevels = fForM.Controls("bx" & iBox)
                
                cVertBarLevels.BackColor = MixColor(sColor, iPrct) 'HexToOleColor(ColorToHex(ColorFromBasePercent(sColor, iPrct)))
                
                iConT = iConT + 1
                iPrct = iConT * 2
            
            Next iBox
            
        End If
    End If
End Sub


'Option Compare Database
'Option Explicit
'
''Initially inspired by http://theaccessweb.com/api/api0060.htm, but the declarations were wrong, so I've reworked it and standardized it.
''Huge help was later found https://www.vbarchiv.net/api/api_choosecolor.html
'
'#If VBA7 Then
'Private Type ChooseColor
'    lStructSize               As Long
'    hwndOwner                 As LongPtr
'    hInstance                 As LongPtr
'    rgbResult                 As Long
'    lpCustColors              As LongPtr
'    flags                     As Long
'    lCustData                 As LongPtr
'    lpfnHook                  As LongPtr
'    lpTemplateName            As String
'End Type
'
'#Else
'
'Private Type ChooseColor
'    lStructSize               As Long
'    hwndOwner                 As Long
'    hInstance                 As Long
'    rgbResult                 As Long
'    lpCustColors              As Long
'    flags                     As Long
'    lCustData                 As Long
'    lpfnHook                  As Long
'    lpTemplateName            As String
'End Type
'
'#End If
'
'Private Const CC_ANYCOLOR = &H100
''Private Const CC_ENABLEHOOK = &H10
''Private Const CC_ENABLETEMPLATE = &H20
''Private Const CC_ENABLETEMPLATEHANDLE = &H40
'Private Const CC_FULLOPEN = &H2
'Private Const CC_PREVENTFULLOPEN = &H4
'Private Const CC_RGBINIT = &H1
''Private Const CC_SHOWHELP = &H8
''Private Const CC_SOLIDCOLOR = &H80
'
'#If VBA7 Then
'    Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'#Else
'    Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'#End If
'
'Public Function ColorPicker(Optional lDefaultColor As Variant) As Long
'    Dim CC                    As ChooseColor
'    Dim lRetVal               As Long
'    Static CustomColors(16)   As Long
'
'    'forma de uso
'    'Call ColorPicker(A), ou
'    'LnG = ColorPicker(A)
'    'onde A, opcional, é a cor que deve vir selecionada por padrão na Cx de Diálogo
'
'    'Some predefined color, there are 16 slots available for predefined colors
'    'You don't have to defined any, if you don't want to!
'    CustomColors(0) = RGB(255, 255, 255)    'White
'    CustomColors(1) = RGB(0, 0, 0)       'Black
'    CustomColors(2) = RGB(255, 0, 0)     'Red
'    CustomColors(3) = RGB(0, 255, 0)     'Green
'    CustomColors(4) = RGB(0, 0, 255)     'Blue
'
'    With CC
'
'        .lStructSize = LenB(CC)
'        .hwndOwner = Application.hWndAccessApp
'        .flags = CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN Or CC_RGBINIT
'        If IsNull(lDefaultColor) = False _
'           And IsMissing(lDefaultColor) = False Then .rgbResult = lDefaultColor    'Set the initial color of the dialog
'        .lpCustColors = VarPtr(CustomColors(0))
'    End With
'    lRetVal = ChooseColor(CC)
'    If lRetVal = 0 Then
'        'Cancelled by the user
'        'ColorPicker = RGB(255, 255, 255)    ' White -> 16777215
'    Else
'        ColorPicker = CC.rgbResult
'    End If
'End Function
'
'
'Function DecToMyColorsRGB_HEX(ByRef lColor As Long) As MyColors
''Converte valor de cor de Decimal, padrão do Access, para valores RGB e para Hexadecimal
'
'    Dim vA, vB
'    Dim lRed As Long, lGreen As Long, lBlue As Long
'    Dim sHex1 As String, sHex2 As String, sHex3 As String
'
'    GetRGB.Red = lColor Mod 256
'    GetRGB.Green = (lColor \ 256) Mod 256
'    GetRGB.Blue = (lColor \ 256 \ 256) Mod 256
'
'    lRed = GetRGB.Red
'    lGreen = GetRGB.Green
'    lBlue = GetRGB.Blue
'
''    sHex1 = Format(CStr(Hex(lRed)), "00")
''    sHex2 = Format(CStr(Hex(lGreen)), "00")
''    sHex3 = Format(CStr(Hex(lBlue)), "00")
'
'    vA = hex(RGB(lBlue, lGreen, lRed))
'    vB = "#" & vA
'
'    DecToMyColorsRGB_HEX.CorAccessGBlng = RGB(lRed, lGreen, lBlue)
'    DecToMyColorsRGB_HEX.CorRGBgbStr = "RGB(" & lRed & "," & lGreen & "," & lBlue & ")"
'    DecToMyColorsRGB_HEX.CorHEXgbStr = "#" & hex(RGB(lBlue, lGreen, lRed))
'    'DecToMyColorsRGB_HEX.CorHEXgbStr = "#" & sHex1 + sHex2 + sHex3
'
'End Function


'---------------------------------------------------------------------
'_TesteColorPicker - orig apagar
'---------------------------------------------------------------------

'Public Sub ColorCheck(cCtrL As Control, cColorValuesLabel As Control)
'
'    Dim vA, vB, vC
'    Dim PickedColor As Long
'    Dim lHEX As Long
'    Dim lngCtlColor As Long
'    Dim LongRGB As Long
'    Dim StringRGB As String
'    Dim StringHEX As String
'    Dim lRed As Long
'    Dim lGreen As Long
'    Dim lBlue As Long
'
'    'Dim cCtrL As Control
'
'    Dim CtlGradientINT As Integer
'    Dim CtlBdColorLNG As Long
'
'    Dim CorPadrao As Long
'
'    Dim CoRhexFunc As MyColors
'
'    Dim CorAccessGotLNG As Long
'    Dim CorStrGotSTR As String
'
'    Dim a, B, c
'    Dim Tmp
'    Dim iPartColor
'
''Stop
'
'    vA = cCtrL.Name
'    iPartColor = MsgBox("Selecione:" & vbCr & vbCr & " - Anular  - para Back color" & vbCr & " - Repetir - para Border color" & vbCr & " - Ignorar - para Fore color", vbAbortRetryIgnore, "Propriedade do controle")
'    '3 Anular
'    '4 Repetir
'    '5 Ignorar
''Stop
'    'If iPartColor = 2 Then Exit Sub
'    lngCtlColor = IIf(iPartColor = 3, cCtrL.BackColor, IIf(iPartColor = 4, cCtrL.BorderColor, cCtrL.ForeColor))
''Stop
'    'a= HSLtoRGB(
'    'HSLtoRGB = rgb(red, green, blue)
''    a = cCtrL.Hue
''    a = cCtrL.Sat
''    a = cCtrL.Lum
'
'    'lngCtlColor = cCtrL.BackColor
'    'lngCtlColor = cCtrL.forecolor
'    'CtlBdColorLNG = cCtrL.BorderColor
'    'CtlGradientINT = cCtrL.Gradient
''Stop
'    PickedColor = ColorPicker(lngCtlColor)
''Stop
'
'    'obtém os valores RGB e HEX a partir do código Long interno do Access
'    CoRes = DecToMyColorsRGB_HEX(PickedColor)
'    'CoRes = DecToMyColorsRGB_HEX(lngCtlColor)
'
'    LongRGB = CoRes.CorAccessGBlng   'longRGB
'    StringRGB = CoRes.CorRGBgbStr      'string RGB
'    StringHEX = CoRes.CorHEXgbStr      'string HEX
'
''Stop
'    cColorValuesLabel.Value = "Cor Access [ " & LongRGB & " ]" & vbCrLf & "Cor RGB [ " & StringRGB & " ]" & vbCrLf & "Cor HEX [ " & StringHEX & " ]"
'    'MsgBox "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
''    Debug.Print "Cor Access [ " & LongRGB & " ]" & vbCr & "Cor RGB [ " & StringRGB & " ]" & vbCr & "Cor HEX [ " & StringHEX & " ]"
'
''Stop
''    GetHSL = MyRGBtoHSL(GetRGB.Red, GetRGB.Green, GetRGB.Blue) ' funcionou
''
''    'Avaliação do HSL
''
''    a = GetHSL.lHue
''    a = GetHSL.lLum
''    If GetHSL.lHue < 11 Or (GetHSL.lHue > 129 And GetHSL.lHue < 240) Then a = "Botão Escuro"
''    If GetHSL.lHue > 21 And GetHSL.lHue < 130 Then a = "Botão Claro"
'
'
''Stop
'    vA = cCtrL.BottomPadding
'    cCtrL.BackColor = LongRGB 'StringRGB 'StringHEX
'    If cCtrL.ControlType = acCommandButton Then cCtrL.Gradient = 25
'End Sub

