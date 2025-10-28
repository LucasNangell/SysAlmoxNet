Attribute VB_Name = "Módulo 00f - Aux (FilterDateFields)"
Option Compare Database
Option Explicit

'WHERE pra filtrar as datas entre a inicial e a final e ainda permitir que a Query seja aberta com o Form fechado
' WHERE (GetDataEntrada1() Is Null   OR DataEntrada >= GetDataEntrada1()) and
' (GetDataEntrada2() Is Null   OR DataEntrada <= GetDataEntrada2())

Public Function GetSrchDataEntrada1() As Variant
Stop
    On Error Resume Next
    If CurrentProject.AllForms("frm_01(1)cProdEstoque").IsLoaded Then
        GetSrchDataEntrada1 = Forms("frm_01(1)cProdEstoque")!txtSrchDataEntrada1
    Else
        GetSrchDataEntrada1 = Null   ' or a default date like #1/1/1900#
    End If
End Function

Public Function GetSrchDataEntrada2() As Variant
Stop
    On Error Resume Next
    If CurrentProject.AllForms("frm_01(1)cProdEstoque").IsLoaded Then
        GetSrchDataEntrada2 = Forms("frm_01(1)cProdEstoque")!txtSrchDataEntrada2
    Else
        GetSrchDataEntrada2 = Null
    End If
End Function
