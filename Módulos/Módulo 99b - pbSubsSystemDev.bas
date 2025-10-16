Attribute VB_Name = "Módulo 99b - pbSubsSystemDev"
Option Compare Database
Option Explicit


'Módulo de apoio para desenvolvimento
' exibe informações sobre as Tabelas do banco


Sub ListarTabelasECampos()
    
'Cria a tabela [ EstruturaBancoDados ]
' .Tabela
' .Campo
' .Tipo de dados
' .Sequência
    
    Dim Db As Database
    Dim tdf As TableDef
    Dim fld As Field
    Dim I As Integer
    Dim strResultado As String
    
    ' Referência ao banco de dados atual
    Set Db = CurrentDb()
    
    ' Cria uma string para armazenar o resultado
    strResultado = "Lista de Tabelas e Campos:" & vbCrLf & vbCrLf
    
    ' Loop através de todas as tabelas
    For Each tdf In Db.TableDefs
        ' Ignora tabelas do sistema (começam com "MSys" ou "~")
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            strResultado = strResultado & "Tabela: " & tdf.Name & vbCrLf
            
            ' Loop através de todos os campos da tabela
            I = 0
            For Each fld In tdf.Fields
                strResultado = strResultado & "   Campo " & I + 1 & ": " & fld.Name & _
                               " (" & TipoCampoParaTexto(fld.Type) & ")" & vbCrLf
                I = I + 1
            Next fld
            
            strResultado = strResultado & vbCrLf
        End If
    Next tdf
    
    ' Exibe o resultado em uma caixa de mensagem (pode ser muito grande para muitas tabelas)
    ' MsgBox strResultado, vbInformation, "Estrutura do Banco de Dados"
    
    ' Melhor alternativa: gravar em um arquivo de texto ou em uma tabela
    GravarResultadoEmArquivo strResultado
    GravarResultadoEmTabela strResultado
    
    ' Limpa objetos
    Set fld = Nothing
    Set tdf = Nothing
    Set Db = Nothing
    
    MsgBox "Listagem concluída! Verifique o arquivo ou tabela de resultados.", vbInformation
End Sub

Function TipoCampoParaTexto(tipo As Integer) As String
    ' Converte o tipo numérico do campo para texto descritivo
    Select Case tipo
        Case dbBoolean: TipoCampoParaTexto = "Sim/Não"
        Case dbByte: TipoCampoParaTexto = "Byte"
        Case dbInteger: TipoCampoParaTexto = "Inteiro"
        Case dbLong: TipoCampoParaTexto = "Longo"
        Case dbCurrency: TipoCampoParaTexto = "Moeda"
        Case dbSingle: TipoCampoParaTexto = "Single"
        Case dbDouble: TipoCampoParaTexto = "Double"
        Case dbDate: TipoCampoParaTexto = "Data/Hora"
        Case dbText: TipoCampoParaTexto = "Texto"
        Case dbLongBinary: TipoCampoParaTexto = "Binário Longo"
        Case dbMemo: TipoCampoParaTexto = "Memo"
        Case dbGUID: TipoCampoParaTexto = "GUID"
        Case Else: TipoCampoParaTexto = "Desconhecido (" & tipo & ")"
    End Select
End Function

Sub GravarResultadoEmArquivo(conteudo As String)
    Dim nomeArquivo As String
    Dim numeroArquivo As Integer
    
    ' Define o nome do arquivo na pasta do banco de dados
    nomeArquivo = CurrentProject.Path & "\EstruturaBancoDados_" & Format(Now(), "yyyy-mm-dd_hh-nn-ss") & ".txt"
    
    ' Obtém um número de arquivo livre
    numeroArquivo = FreeFile()
    
    ' Abre o arquivo para saída
    Open nomeArquivo For Output As #numeroArquivo
    
    ' Escreve o conteúdo
    Print #numeroArquivo, conteudo
    
    ' Fecha o arquivo
    Close #numeroArquivo
End Sub

Sub GravarResultadoEmTabela(conteudo As String)
'    On Error Resume Next
    
    ' Verifica se a tabela de resultados já existe
    Dim tdf As TableDef
    Set tdf = CurrentDb.CreateTableDef("EstruturaBancoDados")
    
    ' Se a tabela não existir, cria
    If Err.Number = 0 Then
        With tdf
            .Fields.Append .CreateField("DataHora", dbDate)
            .Fields.Append .CreateField("Tabela", dbText, 255)
            .Fields.Append .CreateField("Campo", dbText, 255)
            .Fields.Append .CreateField("TipoCampo", dbText, 50)
            .Fields.Append .CreateField("OrdemCampo", dbInteger)
        End With
        CurrentDb.TableDefs.Append tdf
    End If
    
    ' Limpa a tabela se já existir dados
    CurrentDb.Execute "DELETE FROM EstruturaBancoDados", dbFailOnError
    
    ' Insere os dados na tabela
    Dim Db As Database
    Dim tdfSource As TableDef
    Dim fld As Field
    Dim I As Integer
    Dim sql As String
    
    Set Db = CurrentDb()
    
    For Each tdfSource In Db.TableDefs
        If Left(tdfSource.Name, 4) <> "MSys" And Left(tdfSource.Name, 1) <> "~" Then
            I = 0
            For Each fld In tdfSource.Fields
                sql = "INSERT INTO EstruturaBancoDados (DataHora, Tabela, Campo, TipoCampo, OrdemCampo) " & _
                      "VALUES (#" & Now() & "#, '" & Replace(tdfSource.Name, "'", "''") & "', " & _
                      "'" & Replace(fld.Name, "'", "''") & "', '" & TipoCampoParaTexto(fld.Type) & "', " & I + 1 & ")"
                Db.Execute sql, dbFailOnError
                I = I + 1
            Next fld
        End If
    Next tdfSource
    
    ' Limpa objetos
    Set fld = Nothing
    Set tdfSource = Nothing
    Set Db = Nothing
    Set tdf = Nothing
End Sub

Sub ListarControlesDeTodosFormulariosComTag()

'Cria a tabela [ TempControlesComTag ] com as propriedades
' de todos os controles de todos os formulários do banco
' .NomeFormulario
' .NomeControle
' .TipoControle
' .FonteControle
' .Rotulo
' .Tag
' .Visivel
' .Habilitado
' .Topo
' .Esquerda
' .Largura
' .Altura
' .DataHora  (data e hora da criação da tabela)


    If gBbEnableErrorHandler Then On Error Resume Next
    
    Dim accObj As AccessObject
    Dim frm As Form
    Dim CtrL As Control
    Dim strLista As String
    Dim intContador As Integer
    Dim Db As DAO.Database
    Dim rst As DAO.Recordset
    
    Set Db = CurrentDb()
    
    ' Cria uma tabela temporária para armazenar os resultados (se não existir)
    If Not TabelaExiste("TempControlesComTag") Then
        Db.Execute "CREATE TABLE TempControlesComTag (" & _
                   "ID AUTOINCREMENT PRIMARY KEY, " & _
                   "NomeFormulario TEXT(255), " & _
                   "NomeControle TEXT(255), " & _
                   "TipoControle TEXT(255), " & _
                   "FonteControle TEXT(255), " & _
                   "Rotulo TEXT(255), " & _
                   "Tag TEXT(255), " & _
                   "Visivel YESNO, " & _
                   "Habilitado YESNO, " & _
                   "Topo LONG, " & _
                   "Esquerda LONG, " & _
                   "Largura LONG, " & _
                   "Altura LONG, " & _
                   "DataHora DATETIME)"
    Else
        ' Limpa a tabela se já existir
        Db.Execute "DELETE FROM TempControlesComTag"
    End If
    
    Set rst = Db.OpenRecordset("TempControlesComTag")
    
    intContador = 0
    strLista = "LISTA DE CONTROLES DE TODOS OS FORMULÁRIOS (COM TAG)" & vbCrLf & vbCrLf
    
    ' Percorre todos os formulários do projeto
    For Each accObj In CurrentProject.AllForms
        ' Abre o formulário no modo design (invisível)
        DoCmd.OpenForm accObj.Name, acDesign, , , , acHidden
        
        Set frm = Forms(accObj.Name)
        
        strLista = strLista & "FORMULÁRIO: " & accObj.Name & vbCrLf
        
        ' Percorre todos os controles do formulário
        For Each CtrL In frm.Controls
            intContador = intContador + 1
            
            ' Adiciona à string de saída
            strLista = strLista & intContador & ". " & CtrL.Name & " (" & TipoControleParaTexto(CtrL.ControlType) & ")" & vbCrLf
            strLista = strLista & "   Fonte: " & ObterFonteControle(CtrL) & vbCrLf
            strLista = strLista & "   Rótulo: " & ObterRotuloControle(CtrL) & vbCrLf
            strLista = strLista & "   Tag: " & Nz(CtrL.Tag, "(vazia)") & vbCrLf
            strLista = strLista & "   Posição: " & CtrL.Left & ", " & CtrL.Top & vbCrLf
            strLista = strLista & "   Tamanho: " & CtrL.Width & " x " & CtrL.Height & vbCrLf
            strLista = strLista & "   Visível: " & IIf(CtrL.Visible, "Sim", "Não") & vbCrLf
            strLista = strLista & "   Habilitado: " & IIf(CtrL.Enabled, "Sim", "Não") & vbCrLf & vbCrLf
            
            ' Adiciona ao registro na tabela
            With rst
                .AddNew
                !NomeFormulario = accObj.Name
                !NomeControle = CtrL.Name
                !tipoControle = TipoControleParaTexto(CtrL.ControlType)
                !FonteControle = ObterFonteControle(CtrL)
                !Rotulo = ObterRotuloControle(CtrL)
                !Tag = Nz(CtrL.Tag, "(vazia)")
                !Visivel = CtrL.Visible
                !Habilitado = CtrL.Enabled
                !Topo = CtrL.Top
                !Esquerda = CtrL.Left
                !Largura = CtrL.Width
                !Altura = CtrL.Height
                !DataHora = Now()
                .Update
            End With
        Next CtrL
        
        ' Fecha o formulário
        DoCmd.Close acForm, accObj.Name, acSaveNo
        strLista = strLista & "----------------------------------------" & vbCrLf & vbCrLf
    Next accObj
    
    ' Fecha o recordset
    rst.Close
    
    ' Exibe os resultados
    MsgBox "Total de controles encontrados: " & intContador & vbCrLf & vbCrLf & _
           "Os detalhes foram salvos na tabela 'TempControlesComTag'.", vbInformation
    
    ' Opcional: Abre o Immediate Window para ver a lista completa
    If gBbDebugOn Then Debug.Print strLista
    
    ' Opcional: Abre a tabela com os resultados
    DoCmd.OpenTable "TempControlesComTag", acViewNormal, acReadOnly
    
    ' Limpeza
    Set rst = Nothing
    Set Db = Nothing
    Set CtrL = Nothing
    Set frm = Nothing
End Sub

' (Mantenha as mesmas funções auxiliares do código anterior:
' TipoControleParaTexto, ObterFonteControle, ObterRotuloControle, TabelaExiste)

Function TipoControleParaTexto(tipoControle As AcControlType) As String
    ' Converte o tipo numérico do controle para texto descritivo
    Select Case tipoControle
        Case acLabel: TipoControleParaTexto = "Rótulo"
        Case acTextBox: TipoControleParaTexto = "Caixa de Texto"
        Case acComboBox: TipoControleParaTexto = "Caixa de Combinação"
        Case acListBox: TipoControleParaTexto = "Caixa de Listagem"
        Case acCheckBox: TipoControleParaTexto = "Caixa de Seleção"
        Case acOptionButton: TipoControleParaTexto = "Botão de Opção"
        Case acToggleButton: TipoControleParaTexto = "Botão de Alternância"
        Case acCommandButton: TipoControleParaTexto = "Botão de Comando"
        'Case acTabCtrl: TipoControleParaTexto = "Controle de Abas"
        Case acPage: TipoControleParaTexto = "Página"
        Case acPageBreak: TipoControleParaTexto = "Quebra de Página"
        Case acSubform: TipoControleParaTexto = "Subformulário"
        Case acObjectFrame: TipoControleParaTexto = "Quadro de Objeto"
        Case acLine: TipoControleParaTexto = "Linha"
        Case acRectangle: TipoControleParaTexto = "Retângulo"
        Case acImage: TipoControleParaTexto = "Imagem"
        Case acBoundObjectFrame: TipoControleParaTexto = "Quadro de Objeto Vinculado"
        Case acOptionGroup: TipoControleParaTexto = "Grupo de Opções"
        Case acCustomControl: TipoControleParaTexto = "Controle Personalizado"
        Case Else: TipoControleParaTexto = "Desconhecido (" & tipoControle & ")"
    End Select
End Function

Function ObterFonteControle(CtrL As Control) As String
    ' Obtém a fonte de dados do controle (se aplicável)
    If gBbEnableErrorHandler Then On Error Resume Next
    Select Case CtrL.ControlType
        Case acTextBox, acComboBox, acListBox, acCheckBox
            ObterFonteControle = Nz(CtrL.ControlSource, "(não vinculado)")
        Case Else
            ObterFonteControle = "(não aplicável)"
    End Select
End Function

Function ObterRotuloControle(CtrL As Control) As String
    ' Obtém o rótulo associado ao controle (se existir)
    If gBbEnableErrorHandler Then On Error Resume Next
    If Not CtrL.Controls(0) Is Nothing Then
        If CtrL.Controls(0).ControlType = acLabel Then
            ObterRotuloControle = Nz(CtrL.Controls(0).Caption, "(sem rótulo)")
        Else
            ObterRotuloControle = "(sem rótulo)"
        End If
    Else
        ObterRotuloControle = "(sem rótulo)"
    End If
End Function

Function TabelaExiste(nomeTabela As String) As Boolean
    ' Verifica se uma tabela existe no banco de dados atual
    Dim tdf As DAO.TableDef
    TabelaExiste = False
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = nomeTabela Then
            TabelaExiste = True
            Exit For
        End If
    Next tdf
End Function


