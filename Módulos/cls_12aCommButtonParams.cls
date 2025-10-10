VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_12aCommButtonParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public sCtrlName As String  'Nome do controle
Public sFilGrp As String    'Grupo de filtragem a que pertence o controle, no caso de botões de dados
Public sRstArea As String   'Indica a qual [ área de reset ] o botão está associado, no caso de reset area

Public sActType As String   'Indica o tipo de ação que o botão executa
Public sForM As String      'Indica o Formulário que deve ser aberto, no caso de botões de abertura
Public sFrmMode As String   'Indica o FormMode em que o formulário deve ser aberto, no caso de botões de abertura
Public sRecQry As String    'Indica qual consulta será usada para a gravação de dados

Public sOrderMode As String
Public sColToSort As String

