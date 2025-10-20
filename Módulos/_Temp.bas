Attribute VB_Name = "_Temp"
Option Compare Database

Sub testeFoundQryJct()
Dim dict As New Dictionary
Dim qQuery As QueryDef
dict.RemoveAll

For Each qQuery In CurrentDb.QueryDefs
    If InStr(qQuery.sql, "jct") > 0 And InStr(qQuery.sql, "GROUP") = 0 And InStr(qQuery.sql, "qry") = 0 Then
        dict.Add qQuery.Name, qQuery
    End If
Next qQuery
End Sub
