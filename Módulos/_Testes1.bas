Attribute VB_Name = "_Testes1"
Option Compare Database
Option Explicit

'force Listbox update

Private Sub PrSubCtrlChange(vClassCtrl As Variant, sTgtCtrlToUpdt As String)
    Dim vA, vB
    Dim sStR As String
    Dim cCtrL As Control
    Dim fForm As Form
    
'Stop
    Set fForm = vClassCtrl.Parent
    vA = fForm.Name
    
    Set cCtrL = vClassCtrl
    vB = cCtrL.Name
    
    'a rotina Update do próprio controle é chamada antes da rotina AfterUpdate da Classe
    ' por isso é necessário chamar a rotina AfterUpdate do controle ao final da rotina da Classe
'Stop
    sStR = cCtrL.Name & "_AfterUpdate"
    bgbAftUpdtEvntFound = False
'Stop
    If gBbEnableErrorHandler Then On Error Resume Next
    bgbCallingByName = True
    CallByName fForm, sStR, VbMethod   'o evento AfterUpdate deve ser público, do contrário não será localizado
    bgbCallingByName = False
'Stop
    If (Err.Number <> 0) Then
        
        If bgbAftUpdtEvntFound Then
        
            MsgBox "Erro de codificação no evento" & vbCr & "    [ " & sStR & " ] " & vbCr & "do form" & vbCr & "    [ " & fForm.Name & " ] " _
            & vbCr & vbCr & Err.Number & " - " & Err.Description & vbCr & vbCr & "O controle disparador de pesquisa [ " & cCtrL.Name & " ] não irá se comportar como esperado", vbExclamation + vbOKOnly
            
        Else
            MsgBox "O evento" & vbCr & "    [ " & sStR & " ] " & vbCr & "do form" & vbCr & "    [ " & fForm.Name & " ] " & vbCr & vbCr & "- não foi localizado, ou" & vbCr & "- está definido como Privado (Private)" & vbCr _
            & vbCr & vbCr & "O controle disparador de pesquisa [ " & cCtrL.Name & " ] não irá se comportar como esperado", vbExclamation + vbOKOnly
            
        End If
        On Error GoTo -1

'Stop
        Exit Sub
        
    End If
    
    On Error GoTo -1
    
End Sub



'Posteriormente verificar o funcionamento de Nz
Public Sub PbSubSrcKeyPress(ctrForm As Form)

    Dim vA, vB, vC
    Dim vSrcFields As cls_DadosProd
    Set vSrcFields = New cls_DadosProd
Stop
    
    vA = ctrForm.txtSrchProdDescriç
    vA = ctrForm.txtSrchVariaç
    vA = ctrForm.txtSrchCor
    vA = ctrForm.txtSrchMaterial
    vA = ctrForm.txtSrchMed
    vA = ctrForm.txtSrchComp
    
    'vA = ctrForm.cmbSrcCateg.Value, "")
    'vA = ctrForm.cmbSrcAplic.Value, "")
    
    'Armazena na variante [ vSrcFields ] da classe [ cls_DadosProd ],
    ' os valores informados nos campos de filtragem
    ' A classe tem não apenas os campos de filtragem mas também
    '  todos os campos da consulta Fonte de Dados
    vSrcFields.ProdutoDescriç = Nz(ctrForm.txtSrchProdDescriç, "")
    vSrcFields.Variaçao = Nz(ctrForm.txtSrchVariaç, "")
    vSrcFields.ProdCor = Nz(ctrForm.txtSrchCor, "")
    vSrcFields.ProdMaterial = Nz(ctrForm.txtSrchMaterial, "")
    vSrcFields.ProdMedida = Nz(ctrForm.txtSrchMed, "")
    vSrcFields.Complemento = Nz(ctrForm.txtSrchComp, "")
    vSrcFields.ProdCateg = Nz(ctrForm.cmbSrcCateg, "")
    vSrcFields.ProdAplicaçaoDescriç = Nz(ctrForm.cmbSrcAplic, "")
    
    'Prepara o cursor para que o usuário continue fazendo sua digitação
    CampoAtivo.SetFocus
    CampoAtivo.SelStart = Len(CampoAtivo.Text)
    
    Dim sSqlLst$

Stop
    'Chama a função que remonta o SQL da Listbox incluindo os valores de filtragem
    sSqlLst = sqlCreate(vSrcFields, ctrForm.lstProdsBase)

Stop
        
    ctrForm.lstProdsBase.RowSource = sSqlLst

    
End Sub




'        'O evento deve ser público, do contrário não será localizado
'        sCtrlEvent = cCtrL.Name & "_Change"
'        Debug.Print sCtrlEvent
'
''        On Error Resume Next
'        gBbCallingByName = True
'        CallByName fForM, sCtrlEvent, VbMethod   'o evento AfterUpdate deve ser público, do contrário não será localizado
'        gBbEventFound = True
'        If (Err.Number = 2465) Then 'Evento não existe ou não está Público
'            gBbEventFound = False
'        End If
'        On Error GoTo -1
'
''        gBbCallingByName = True
''        CallByName fForM, sStR, VbMethod   'o evento AfterUpdate deve ser público, do contrário não será localizado
''        gBbCallingByName = False
''        If (Err.Number <> 0) Then
''
''            If bgbAftUpdtEvntFound Then
''
''                MsgBox "Erro de codificação no evento" & vbCr & "    [ " & sStR & " ] " & vbCr & "do form" & vbCr & "    [ " & fForM.Name & " ] " _
''                & vbCr & vbCr & Err.Number & " - " & Err.Description & vbCr & vbCr & "O controle disparador de pesquisa [ " & cTriggCtl.Name & " ] não irá se comportar como esperado", vbExclamation + vbOKOnly
''
''            Else
''                MsgBox "O evento" & vbCr & "    [ " & sStR & " ] " & vbCr & "do form" & vbCr & "    [ " & fForM.Name & " ] " & vbCr & vbCr & "- não foi localizado, ou" & vbCr & "- está definido como Privado (Private)" & vbCr _
''                & vbCr & vbCr & "O controle disparador de pesquisa [ " & cTriggCtl.Name & " ] não irá se comportar como esperado", vbExclamation + vbOKOnly
''
''            End If
''            On Error GoTo -1
''
'''Stop
''            Exit Sub
''
''        End If
''        On Error GoTo -1


Sub Teste()

     For Each cCtrL In clObjFormOpenParams.fTrgtForm
         
         If InStr(cCtrL.Tag, "FormMode") > 0 Then
             
             If Not IsObject(dictCtrlsEvents(clObjFormOpenParams.sTrgtForm)) Then Set dictCtrlsEvents(clObjFormOpenParams.sTrgtForm) = New Dictionary
             
             If Not dictCtrlsEvents(clObjFormOpenParams.sTrgtForm).Exists(cCtrL.Name) = True Then
        
                'Cria um novo objeto [ clObjCtrlsEvents ] da Classe [ cls_10aCtrls_Events ] pra ser incluído no [ dictCtrlsEvents(sForM) ]
                 Set clObjCtrlsEvents = New cls_10aCtrls_Events
                 dictCtrlsEvents(clObjFormOpenParams.sTrgtForm).Add cCtrL.Name, clObjCtrlsEvents
                 
                 clObjCtrlsEvents.sCtrlName = cCtrL.Name
                 'A inicialização dos controles será feita posteriormente, na abertura do formulário
             
             End If
        End If
    
    Next cCtrL

End Sub

