Attribute VB_Name = "Módulo 99c -  pbSubsSystemDevGit"
Option Compare Database
' Adicionar esta declaração no topo do módulo (se ainda não existir)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub ExportarTodosObjetosVBADetallado()
    Dim vbComp As VBIDE.VBComponent
    Dim strFolderPath As String
    Dim strExtension As String
    Dim strTipoObjeto As String
    Dim intContador As Integer
    Dim qry As QueryDef
    Dim tdf As TableDef
    Dim obj As AccessObject
    Dim x
    
    ' Especifica la ruta de la carpeta destino
    strFolderPath = CurrentProject.Path & "\Módulos\"
    
    ' Verificar si la carpeta existe, si no crearla
    If Dir(strFolderPath, vbDirectory) = "" Then
        MkDir strFolderPath
    End If
    
    ' Crear subcarpetas para organizar
    If Dir(strFolderPath & "forms\", vbDirectory) = "" Then
        MkDir strFolderPath & "forms\"
    End If
    If Dir(strFolderPath & "queries\", vbDirectory) = "" Then
        MkDir strFolderPath & "queries\"
    End If
    If Dir(strFolderPath & "tables\", vbDirectory) = "" Then
        MkDir strFolderPath & "tables\"
    End If
    If Dir(strFolderPath & "reports\", vbDirectory) = "" Then
        MkDir strFolderPath & "reports\"
    End If
    If Dir(strFolderPath & "macros\", vbDirectory) = "" Then
        MkDir strFolderPath & "macros\"
    End If
    If Dir(strFolderPath & "forms_xml\", vbDirectory) = "" Then
        MkDir strFolderPath & "forms_xml\"
    End If
    If Dir(strFolderPath & "reports_xml\", vbDirectory) = "" Then
        MkDir strFolderPath & "reports_xml\"
    End If
    
    intContador = 0

    ' --- EXPORTAR CÓDIGO VBA ---
    For Each vbComp In Application.VBE.VBProjects(1).VBComponents
        'debug.print vbComp.Type & "| " & vbComp.Name
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                strExtension = ".bas"
                strTipoObjeto = "Módulo"
            Case vbext_ct_ClassModule
                strExtension = ".cls"
                strTipoObjeto = "Clase"
            Case 100
                strExtension = ".cls"
                strTipoObjeto = "Formulario"
            Case Else
                strExtension = ""
                strTipoObjeto = "Otro"
        End Select
        
        If strExtension <> "" Then
            On Error Resume Next
            vbComp.Export strFolderPath & vbComp.Name & strExtension
            If Err.Number = 0 Then
                intContador = intContador + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next vbComp

    ' --- EXPORTAR FORMULÁRIOS (formato legível para Git) ---
    For Each obj In CurrentProject.AllForms
        On Error Resume Next
        
        ' Exportar em formato XML (Git-friendly)
        Call ExportarFormularioParaXML(obj.Name, strFolderPath & "forms_xml\" & obj.Name & ".xml")
        
        ' Exportar também em formato padrão para backup
        Application.SaveAsText acForm, obj.Name, strFolderPath & "forms\" & obj.Name & ".form"
        
        If Err.Number = 0 Then
            intContador = intContador + 1
            'debug.print "Formulário exportado: " & obj.Name
        Else
            'debug.print "Erro ao exportar formulário " & obj.Name & ": " & Err.Description
            Err.Clear
        End If
        DoCmd.Close acForm, obj.Name
        
        On Error GoTo 0
    Next obj

    ' --- EXPORTAR RELATÓRIOS (formato legível para Git) ---
    For Each obj In CurrentProject.AllReports
        On Error Resume Next
        
        ' Exportar em formato XML (Git-friendly)
        Call ExportarRelatorioParaXML(obj.Name, strFolderPath & "reports_xml\" & obj.Name & ".xml")
        
        ' Exportar também em formato padrão para backup
        Application.SaveAsText acReport, obj.Name, strFolderPath & "reports\" & obj.Name & ".report"
        
        If Err.Number = 0 Then
            intContador = intContador + 1
            'debug.print "Relatório exportado: " & obj.Name
        Else
            'debug.print "Erro ao exportar relatório " & obj.Name & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next obj

    ' --- EXPORTAR CONSULTAS (SQL) ---
    For Each qry In CurrentDb.QueryDefs
        ' Ignorar consultas temporárias e de sistema
        If Left(qry.Name, 1) <> "~" And qry.Name <> "~sq_ck" Then
            On Error Resume Next
            Call ExportarQueryParaSQL(qry, strFolderPath & "queries\" & qry.Name & ".sql")
            If Err.Number = 0 Then
                intContador = intContador + 1
                'debug.print "Consulta exportada: " & qry.Name
            Else
                'debug.print "Erro ao exportar consulta " & qry.Name & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next qry

    ' --- EXPORTAR TABELAS (estrutura) ---
    For Each tdf In CurrentDb.TableDefs
        ' Ignorar tabelas do sistema
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            On Error Resume Next
            Call ExportarTabelaParaSQL(tdf, strFolderPath & "tables\" & tdf.Name & ".sql")
            If Err.Number = 0 Then
                intContador = intContador + 1
                'debug.print "Tabela exportada: " & tdf.Name
            Else
                'debug.print "Erro ao exportar tabela " & tdf.Name & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next tdf

    ' --- EXPORTAR MACROS ---
    For Each obj In CurrentProject.AllMacros
        On Error Resume Next
        Application.SaveAsText acMacro, obj.Name, strFolderPath & "macros\" & obj.Name & ".macro"
        If Err.Number = 0 Then
            intContador = intContador + 1
            'debug.print "Macro exportada: " & obj.Name
        Else
            'debug.print "Erro ao exportar macro " & obj.Name & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next obj

    ' --- COMANDOS GIT ---
    Dim strComand As String
    strComand = InputBox("Digite a mensagem da versão")


    x = Shell("cmd.exe /K ""cd /d " & CurrentProject.Path, 1)
    Sleep (500)
    
    SendKeys ("git add *")
    SendKeys ("{ENTER}")
    Sleep (200)
    SendKeys ("git commit -m " & """" & strComand & """")
    SendKeys ("{ENTER}")
    Sleep (200)
    SendKeys ("git push")
    SendKeys ("{ENTER}")
    SendKeys ("exit")
    SendKeys ("{ENTER}")
    
    'MsgBox "Exportação e commit concluídos! " & intContador & " objetos exportados.", vbInformation
End Sub

' --- NOVAS FUNÇÕES CORRIGIDAS PARA EXPORTAR EM XML ---

' Função para exportar formulário em XML legível (CORRIGIDA)
Sub ExportarFormularioParaXML(strFormName As String, strFilePath As String)
    On Error GoTo ErroHandler
    
    Dim frm As Form
    Dim ctl As Control
    Dim intFile As Integer
    Dim strXML As String
    Dim blnFormAberto As Boolean
    
    blnFormAberto = False
    
    ' Verificar se o formulário já está aberto
    If CurrentProject.AllForms(strFormName).IsLoaded Then
        ' Se já está aberto, usar a instância existente
        Set frm = Forms(strFormName)
        blnFormAberto = True
    Else
        ' Abrir o formulário em modo design (oculto)
        DoCmd.OpenForm strFormName, acDesign, , , , acHidden
        Set frm = Forms(strFormName)
        blnFormAberto = True
    End If
    
    ' Iniciar XML
    strXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strXML = strXML & "<Form name=""" & strFormName & """>" & vbCrLf
    
    ' Propriedades do formulário
    strXML = strXML & "  <Properties>" & vbCrLf
    strXML = strXML & "    <Caption>" & SafeXMLEncode(GetProperty(frm, "Caption")) & "</Caption>" & vbCrLf
    strXML = strXML & "    <RecordSource>" & SafeXMLEncode(GetProperty(frm, "RecordSource")) & "</RecordSource>" & vbCrLf
    strXML = strXML & "    <Width>" & GetProperty(frm, "Width") & "</Width>" & vbCrLf
    strXML = strXML & "    <Height>" & GetProperty(frm, "Height") & "</Height>" & vbCrLf
    strXML = strXML & "    <DefaultView>" & GetProperty(frm, "DefaultView") & "</DefaultView>" & vbCrLf
    strXML = strXML & "  </Properties>" & vbCrLf
    
    ' Controles
    strXML = strXML & "  <Controls>" & vbCrLf
    
    For Each ctl In frm.Controls
        On Error Resume Next
        strXML = strXML & "    <Control>" & vbCrLf
        strXML = strXML & "      <Name>" & SafeXMLEncode(GetControlProperty(ctl, "Name")) & "</Name>" & vbCrLf
        strXML = strXML & "      <Type>" & TypeName(ctl) & "</Type>" & vbCrLf
        strXML = strXML & "      <Left>" & GetControlProperty(ctl, "Left") & "</Left>" & vbCrLf
        strXML = strXML & "      <Top>" & GetControlProperty(ctl, "Top") & "</Top>" & vbCrLf
        strXML = strXML & "      <Width>" & GetControlProperty(ctl, "Width") & "</Width>" & vbCrLf
        strXML = strXML & "      <Height>" & GetControlProperty(ctl, "Height") & "</Height>" & vbCrLf
        strXML = strXML & "      <Tag>" & GetProperty(ctl, "Tag") & "</Tag>" & vbCrLf
        strXML = strXML & "      <Enabled>" & GetProperty(ctl, "Enabled") & "</Enabled>" & vbCrLf
        ' Propriedades específicas por tipo de controle
        Select Case TypeName(ctl)
            Case "TextBox"
                strXML = strXML & "      <ControlSource>" & SafeXMLEncode(GetControlProperty(ctl, "ControlSource")) & "</ControlSource>" & vbCrLf
                strXML = strXML & "      <Format>" & SafeXMLEncode(GetControlProperty(ctl, "Format")) & "</Format>" & vbCrLf
            Case "Label"
                strXML = strXML & "      <Caption>" & SafeXMLEncode(GetControlProperty(ctl, "Caption")) & "</Caption>" & vbCrLf
            Case "ComboBox", "ListBox"
                strXML = strXML & "      <RowSource>" & SafeXMLEncode(GetControlProperty(ctl, "RowSource")) & "</RowSource>" & vbCrLf
            Case "CommandButton"
                strXML = strXML & "      <Caption>" & SafeXMLEncode(GetControlProperty(ctl, "Caption")) & "</Caption>" & vbCrLf
                strXML = strXML & "      <OnClick>" & SafeXMLEncode(GetControlProperty(ctl, "OnClick")) & "</OnClick>" & vbCrLf
        End Select
        
        strXML = strXML & "    </Control>" & vbCrLf
        On Error GoTo 0
    Next ctl
    
    strXML = strXML & "  </Controls>" & vbCrLf
    strXML = strXML & "</Form>"
    
    ' Salvar arquivo
    intFile = FreeFile
    Open strFilePath For Output As intFile
    Print #intFile, strXML
    Close intFile
    
    ' Fechar formulário apenas se nós o abrimos
    If blnFormAberto And Not CurrentProject.AllForms(strFormName).IsLoaded Then
        DoCmd.Close acForm, strFormName, acSaveNo
    End If
    
    Exit Sub

ErroHandler:
    ' Tentar fechar o formulário em caso de erro
    If blnFormAberto Then
        On Error Resume Next
        DoCmd.Close acForm, strFormName, acSaveNo
        On Error GoTo 0
    End If
    
    ' Registrar erro
    Err.Raise Err.Number, "ExportarFormularioParaXML", "Erro ao exportar formulário " & strFormName & ": " & Err.Description
End Sub

' Função para exportar relatório em XML legível (CORRIGIDA)
Sub ExportarRelatorioParaXML(strReportName As String, strFilePath As String)
    On Error GoTo ErroHandler
    
    Dim rpt As Report
    Dim ctl As Control
    Dim intFile As Integer
    Dim strXML As String
    Dim blnReportAberto As Boolean
    
    blnReportAberto = False
    
    ' Abrir o relatório em modo design (oculto)
    DoCmd.OpenReport strReportName, acViewDesign, , , , acHidden
    Set rpt = Reports(strReportName)
    blnReportAberto = True
    
    ' Iniciar XML
    strXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strXML = strXML & "<Report name=""" & strReportName & """>" & vbCrLf
    
    ' Propriedades do relatório
    strXML = strXML & "  <Properties>" & vbCrLf
    strXML = strXML & "    <Caption>" & SafeXMLEncode(GetProperty(rpt, "Caption")) & "</Caption>" & vbCrLf
    strXML = strXML & "    <RecordSource>" & SafeXMLEncode(GetProperty(rpt, "RecordSource")) & "</RecordSource>" & vbCrLf
    strXML = strXML & "    <Width>" & GetProperty(rpt, "Width") & "</Width>" & vbCrLf
    strXML = strXML & "    <Height>" & GetProperty(rpt, "Height") & "</Height>" & vbCrLf
    strXML = strXML & "  </Properties>" & vbCrLf
    
    ' Controles
    strXML = strXML & "  <Controls>" & vbCrLf
    
    For Each ctl In rpt.Controls
        On Error Resume Next
        strXML = strXML & "    <Control>" & vbCrLf
        strXML = strXML & "      <Name>" & SafeXMLEncode(GetControlProperty(ctl, "Name")) & "</Name>" & vbCrLf
        strXML = strXML & "      <Type>" & TypeName(ctl) & "</Type>" & vbCrLf
        strXML = strXML & "      <Left>" & GetControlProperty(ctl, "Left") & "</Left>" & vbCrLf
        strXML = strXML & "      <Top>" & GetControlProperty(ctl, "Top") & "</Top>" & vbCrLf
        strXML = strXML & "      <Width>" & GetControlProperty(ctl, "Width") & "</Width>" & vbCrLf
        strXML = strXML & "      <Height>" & GetControlProperty(ctl, "Height") & "</Height>" & vbCrLf
        
        ' Propriedades específicas por tipo de controle
        Select Case TypeName(ctl)
            Case "TextBox"
                strXML = strXML & "      <ControlSource>" & SafeXMLEncode(GetControlProperty(ctl, "ControlSource")) & "</ControlSource>" & vbCrLf
                strXML = strXML & "      <Format>" & SafeXMLEncode(GetControlProperty(ctl, "Format")) & "</Format>" & vbCrLf
            Case "Label"
                strXML = strXML & "      <Caption>" & SafeXMLEncode(GetControlProperty(ctl, "Caption")) & "</Caption>" & vbCrLf
        End Select
        
        strXML = strXML & "    </Control>" & vbCrLf
        On Error GoTo 0
    Next ctl
    
    strXML = strXML & "  </Controls>" & vbCrLf
    strXML = strXML & "</Report>"
    
    ' Salvar arquivo
    intFile = FreeFile
    Open strFilePath For Output As intFile
    Print #intFile, strXML
    Close intFile
    
    ' Fechar relatório
    DoCmd.Close acReport, strReportName, acSaveNo
    blnReportAberto = False
    
    Exit Sub

ErroHandler:
    ' Tentar fechar o relatório em caso de erro
    If blnReportAberto Then
        On Error Resume Next
        DoCmd.Close acReport, strReportName, acSaveNo
        On Error GoTo 0
    End If
    
    ' Registrar erro
    Err.Raise Err.Number, "ExportarRelatorioParaXML", "Erro ao exportar relatório " & strReportName & ": " & Err.Description
End Sub

' --- FUNÇÕES AUXILIARES SEGURAS ---

' Função segura para obter propriedades de objetos
Function GetProperty(obj As Object, propName As String) As Variant
    On Error Resume Next
    GetProperty = CallByName(obj, propName, VbGet)
    If Err.Number <> 0 Then
        GetProperty = ""
        Err.Clear
    End If
End Function

' Função segura para obter propriedades de controles
Function GetControlProperty(ctl As Control, propName As String) As Variant
    On Error Resume Next
    GetControlProperty = CallByName(ctl, propName, VbGet)
    If Err.Number <> 0 Then
        GetControlProperty = ""
        Err.Clear
    End If
End Function

' Função segura para codificar caracteres especiais em XML
Function SafeXMLEncode(varText As Variant) As String
    On Error Resume Next
    
    If IsNull(varText) Then
        SafeXMLEncode = ""
        Exit Function
    End If
    
    Dim strText As String
    strText = CStr(varText)
    
    SafeXMLEncode = Replace(strText, "&", "&amp;")
    SafeXMLEncode = Replace(SafeXMLEncode, "<", "&lt;")
    SafeXMLEncode = Replace(SafeXMLEncode, ">", "&gt;")
    SafeXMLEncode = Replace(SafeXMLEncode, """", "&quot;")
    SafeXMLEncode = Replace(SafeXMLEncode, "'", "&apos;")
End Function

' --- FUNÇÕES AUXILIARES EXISTENTES ---

' Função para exportar consultas como SQL
Sub ExportarQueryParaSQL(qry As QueryDef, strCaminhoCompleto As String)
    Dim intArquivo As Integer
    Dim strSQL As String
    Dim strTipo As String
    
    ' Determinar tipo da consulta
    Select Case qry.Type
        Case 0: strTipo = "SELECT"
        Case 1: strTipo = "UNION"
        Case 2: strTipo = "CROSSTAB"
        Case 3: strTipo = "MAKE_TABLE"
        Case 4: strTipo = "APPEND"
        Case 5: strTipo = "UPDATE"
        Case 6: strTipo = "DELETE"
        Case 7: strTipo = "DDL"
        Case Else: strTipo = "UNKNOWN"
    End Select
    
    ' Criar cabeçalho com metadados
    strSQL = "-- Consulta: " & qry.Name & vbCrLf
    strSQL = strSQL & "-- Tipo: " & strTipo & vbCrLf
    'strSQL = strSQL & "-- Exportado em: " & Now() & vbCrLf
    strSQL = strSQL & "-- Banco: " & CurrentProject.Name & vbCrLf & vbCrLf
    
    ' Adicionar o SQL original
    strSQL = strSQL & qry.sql
    
    ' Salvar arquivo
    intArquivo = FreeFile
    Open strCaminhoCompleto For Output As intArquivo
    Print #intArquivo, strSQL
    Close intArquivo
End Sub

' Função para exportar estrutura de tabelas como SQL
Sub ExportarTabelaParaSQL(tdf As TableDef, strCaminhoCompleto As String)
    Dim intArquivo As Integer
    Dim strSQL As String
    Dim fld As Field
    Dim idx As Index
    Dim strCampos As String
    Dim strChaves As String
    
    ' Cabeçalho
    strSQL = "-- Tabela: " & tdf.Name & vbCrLf
    'strSQL = strSQL & "-- Exportado em: " & Now() & vbCrLf
    strSQL = strSQL & "-- Registros: " & DCount("*", tdf.Name) & vbCrLf & vbCrLf
    
    ' Criar CREATE TABLE
    strSQL = strSQL & "CREATE TABLE " & tdf.Name & " (" & vbCrLf
    
    ' Campos
    For Each fld In tdf.Fields
        strCampos = strCampos & "    " & fld.Name & " " & TipoDadosParaSQL(fld.Type)
        
        ' Adicionar propriedades do campo
        If fld.Required Then strCampos = strCampos & " NOT NULL"
        If fld.DefaultValue <> "" Then strCampos = strCampos & " DEFAULT " & fld.DefaultValue
        If fld.Attributes And dbAutoIncrField Then strCampos = strCampos & " AUTOINCREMENT"
        
        strCampos = strCampos & "," & vbCrLf
    Next fld
    
    ' Remover última vírgula
    If Len(strCampos) > 0 Then
        strCampos = Left(strCampos, Len(strCampos) - 3) & vbCrLf
    End If
    
    strSQL = strSQL & strCampos & ");" & vbCrLf & vbCrLf
    
    ' Índices e chaves primárias
    For Each idx In tdf.Indexes
        If idx.Primary Then
            strSQL = strSQL & "ALTER TABLE " & tdf.Name & " ADD PRIMARY KEY ("
        ElseIf idx.Unique Then
            strSQL = strSQL & "CREATE UNIQUE INDEX " & idx.Name & " ON " & tdf.Name & " ("
        Else
            strSQL = strSQL & "CREATE INDEX " & idx.Name & " ON " & tdf.Name & " ("
        End If
        
        Dim fldIdx As Field
        For Each fldIdx In idx.Fields
            strChaves = strChaves & fldIdx.Name & ", "
        Next fldIdx
        
        If Len(strChaves) > 0 Then
            strChaves = Left(strChaves, Len(strChaves) - 2)
        End If
        
        strSQL = strSQL & strChaves & ");" & vbCrLf
        strChaves = ""
    Next idx
    
    ' Salvar arquivo
    intArquivo = FreeFile
    Open strCaminhoCompleto For Output As intArquivo
    Print #intArquivo, strSQL
    Close intArquivo
End Sub

' Função auxiliar para converter tipos de dados do Access para SQL
Function TipoDadosParaSQL(intTipo As Integer) As String
    Select Case intTipo
        Case dbBoolean: TipoDadosParaSQL = "BIT"
        Case dbByte: TipoDadosParaSQL = "TINYINT"
        Case dbInteger: TipoDadosParaSQL = "SMALLINT"
        Case dbLong: TipoDadosParaSQL = "INTEGER"
        Case dbCurrency: TipoDadosParaSQL = "MONEY"
        Case dbSingle: TipoDadosParaSQL = "REAL"
        Case dbDouble: TipoDadosParaSQL = "FLOAT"
        Case dbDate: TipoDadosParaSQL = "DATETIME"
        Case dbText: TipoDadosParaSQL = "VARCHAR(" & 255 & ")" ' Acess usa 255 como padrão
        Case dbLongBinary: TipoDadosParaSQL = "VARBINARY(MAX)"
        Case dbMemo: TipoDadosParaSQL = "TEXT"
        Case dbGUID: TipoDadosParaSQL = "UNIQUEIDENTIFIER"
        Case Else: TipoDadosParaSQL = "VARCHAR(255)"
    End Select
End Function



