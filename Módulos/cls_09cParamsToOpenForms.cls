VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_09cParamsToOpenForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public sTrgtForm As String
Public sTrggForM As String
Public sFormMode As String
Public sFilGrp As String
Public fTrgtForm As Form
Public bSetPosition As Boolean
Public bFrmIsDocked As Boolean
Public bCentralizeForm As Boolean
Public bIsTrggForm As Boolean
Public bShowRecID As Boolean
Public cTglBtnDocking As Control
Public lngRecID As Long
Public lngFormLeft As Long
Public lngFormTop As Long
Public vOpenArgs As Variant

Public cTrgtCtrl As Control
Public sQryFieldID As String
