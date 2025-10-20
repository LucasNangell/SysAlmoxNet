Attribute VB_Name = "_Testes4 BuildDictCtrlsInBox"
Option Compare Database
Option Explicit

Type PosCtrl
    Left As Long
    Top As Long
    Right As Long
    Botton As Long
End Type
Public tPosCtrl As PosCtrl

Public Sub DictCtrlsInBoxBuild(cBox As Control)
    Dim cCtrL As Control
    Dim sForM As String
    
    sForM = cBox.Parent.Name
    
    If Not IsObject(dictCtrlsInBox(sForM)) Then Set dictCtrlsInBox(sForM) = New Dictionary
    If Not IsObject(dictCtrlsInBox(sForM)(cBox.Name)) Then Set dictCtrlsInBox(sForM)(cBox.Name) = New Dictionary
    
    
    For Each cCtrL In cBox.Parent
        If CtrlIsInBox(cCtrL, cBox) Then
            If InStr(cCtrL.Tag, "NotInDict") = 0 Then
                Select Case cCtrL.ControlType
                    Case acTextBox, acComboBox, acListBox, acCommandButton, acToggleButton, acOptionButton, acCheckBox
                        
                        If Not dictCtrlsInBox(sForM)(cBox.Name).Exists(cCtrL.Name) Then
                            dictCtrlsInBox(sForM)(cBox.Name).Add cCtrL.Name, cCtrL
                            'Debug.Print cCtrL.Name
                        End If
                End Select
            End If
        End If
    Next cCtrL
    
End Sub
Public Function GetPosCtrl(cCtrL As Control) As PosCtrl
    GetPosCtrl.Left = cCtrL.Left
    GetPosCtrl.Top = cCtrL.Top
    GetPosCtrl.Right = cCtrL.Left + cCtrL.Width
    GetPosCtrl.Botton = cCtrL.Top + cCtrL.Height
End Function
Public Function CtrlIsInBox(cCtrL As Control, cBox As Control) As Boolean
    Dim tBoxPosition As PosCtrl
    Dim tCtrlPosition As PosCtrl
    
    tBoxPosition = GetPosCtrl(cBox)
    tCtrlPosition = GetPosCtrl(cCtrL)
    
    If tCtrlPosition.Left > tBoxPosition.Left And _
       tCtrlPosition.Top > tBoxPosition.Top And _
       tCtrlPosition.Right < tBoxPosition.Right And _
       tCtrlPosition.Botton < tBoxPosition.Botton Then _
       CtrlIsInBox = True Else CtrlIsInBox = False

End Function



