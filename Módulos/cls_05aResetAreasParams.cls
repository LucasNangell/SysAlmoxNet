VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_05aResetAreasParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public sRstAr As String     'Indica a área de reset do controle

Public dictRstArCtrls As New Dictionary  'Dicionário pra guardar todos os [ controles ] que devem ser esvaziados
Public dictRstArFilGrps As New Dictionary    'Dicionário pra guardar todos os [ grupos de filtragem ] a serem atualizados


 
 
