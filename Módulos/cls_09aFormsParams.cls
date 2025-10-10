VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_09aFormsParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public sTrggFormName As String  'Nome do formulário disparador
Public sTrgtFormName As String  'Nome do formulário alvo

'Coordenadas do TriggerForm
Public lngTrggLeft As Long
Public lngTrggTop As Long
Public lngTrggWidth As Long

'Coordenadas do TargtForm

Public lngTrgtLeft As Long
Public lngTrgtTop As Long

Public lngTrgtHorizOffset As Long
Public lngTrgtVertOffset As Long

Public bTrgtFormIsDocked As Boolean


'Informações a exibir na Barra de Status do Form
'-----------------------------------------------
Public bForm_ShowWarns As Boolean
'Public sForm_StatusBarText As String

' -->  dicionário, dentro do objeto de classe pra armazenar alertas a serem exibidos na Barra de Status do Form
Public dForm_StatusBarWarns As New Dictionary



'Public s_FormNoTrgtCtrls As String

'Public s_MissingHLclrParams As String
'Public dictHLclrWarn As Dictionary




    
