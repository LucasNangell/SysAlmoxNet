VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_07bLckdStatusParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Parâmetros de Controle que alternam entre HABILITADO e DESABILITADO  ("Enbld", "Dsbld"}
Public lngLckdStatusBackColor As Long        'cor de fundo
Public lngLckdStatusForeColor As Long        'cor de primeiro plano
Public lngLckdStatusBorderColor As Long      'cor da borda
Public lngLckdStatusBorderStyle As Long      'estilo da borda (Sólida/Transp., exceto para botões)
Public iLckdStatusSpecialEffect As Integer   'aparência do controle (relevo, embossed, etc, exceto para Commandbuttons)



