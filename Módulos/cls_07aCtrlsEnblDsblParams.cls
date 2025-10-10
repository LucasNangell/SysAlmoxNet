VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_07aCtrlsEnblDsblParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'-----------------------------------------------------------------------
'Parâmetros de aparência dos controles com base no status Enable/Disable
'-----------------------------------------------------------------------

Public sCtrlName As String
'Dim sFormName As String
Dim sTipText As String
'Dim bEnable As Boolean
'Dim bVisible As Boolean

'Parâmetros comuns a Enabled e Disabled
Public sInitTipText As String    'texto original da Dica de Controle
Public iCtrlGradientColor As Integer 'cor do gradiente (apenas para Commandbuttons)

Public dictParamByLckdStatus As Dictionary



''Parâmetros do Controle HABILITADO
' Public lngEnbldCtrlBackColor As Long        'cor de fundo
' Public lngEnbldCtrlForeColor As Long        'cor de primeiro plano
' Public lngEnbldCtrlBorderColor As Long      'cor da borda
' Public lngEnbldCtrlBorderStyle As Long      'estilo da borda (Sólida/Transp., exceto para botões)
' Public iEnbldCtrlSpecialEffect As Integer   'aparência do controle (relevo, embossed, etc, exceto para Commandbuttons)

''Parâmetros do Controle DESABILITADO
' Public lngDsbldCtrlBackColor As Long        'cor de fundo
' Public lngDsbldCtrlForeColor As Long        'cor de primeiro plano
' Public lngDsbldCtrlBorderColor As Long      'cor da borda
' Public lngDsbldCtrlBorderStyle As Long      'estilo da borda (Sólida/Transp., exceto para botões)
' Public iDsbldCtrlSpecialEffect As Integer   'aparência do controle (relevo, embossed, etc, exceto para Commandbuttons)



