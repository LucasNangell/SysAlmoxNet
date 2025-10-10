VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_11aCtrlBehvrParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public sCtrlName As String  'nome do controle

'-----------------------------------------------------------------------
'Parâmetros de comportamento dos controles
'-----------------------------------------------------------------------

Public bColorHighlight As Boolean   'indica se o controle deve mudar de cor no foco
Public bOnDirty As Boolean          'indica se o controle deve mudar de cor no sujo (apenas para Triggers)
                                    ' (se pertencer a um Grupo de Filtragem ou se for um TgtCtrl com indicação de mudança de cor no Sujo)
Public bMskdCtrl As Boolean       'indica se o controle (texbox) tem máscara de exibição a ser desconsiderada na filtragem embora exibida no controle


Public bMskdCtrlEventFound As Boolean  'a rotina [ MskdTxtbox02_TextMask ] só deve ser ativada se o controle tiver simultaneamente:
                                       ' . um evento Change com a rotina MskdTxtbox02_TextMask nele incluída
                                       ' . no [ dictCtrlBehvrParams(sForM) ] o parâmetro [ bMskd ] TRUE


'Public bTriggrddByCtrlEvent As Boolean  'confirma que a filtragem do controle tenha sido disparada pela Rotina MskdTxtbox02_TextMask
                                        ' Isso evita que um controle indevidamente configurado como bMskd = TRUE tenha sua filtragem alterada




