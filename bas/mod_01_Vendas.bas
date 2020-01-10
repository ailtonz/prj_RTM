Attribute VB_Name = "mod_01_Vendas"
Option Base 1
Option Explicit

Public Function CadastroEspecial(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroEspecial = True
On Error GoTo CadastroEspecial_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroEspecial As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroEspecial = dbOrcamento.QueryDefs("CadastroEspecial")

With qdfCadastroEspecial

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    'QUANTIDADE
    l = 9
    c = 3
    For x = 1 To 8
        .Parameters(x & "QTD") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'FORMATO-1
    .Parameters("1FORMATO") = Cells(11, 3).Value
    .Parameters("2FORMATO") = Cells(11, 7).Value
    
    'FORMATO-2
    .Parameters("3FORMATO") = Cells(12, 3).Value
    .Parameters("4FORMATO") = Cells(12, 7).Value
    
    
    'DESCRIÇÃO
    l = 15
    c = 2
    For x = 1 To 4
        .Parameters(x & "DESCRICAO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'N° DE PÁGINAS
    l = 15
    c = 3
    For x = 1 To 4
        .Parameters(x & "NPAGINAS") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'CORES
    l = 15
    c = 5
    For x = 1 To 4
        .Parameters(x & "CORES") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'PAPEL
    l = 15
    c = 7
    For x = 1 To 4
        .Parameters(x & "PAPEL") = Cells(l, c).Value
        l = l + 1
    Next x
    
    .Parameters("1ACABAMENTO") = Range("C19").Value
    .Parameters("2ACABAMENTO") = Range("C20").Value
    
    .Execute
    
End With

qdfCadastroEspecial.Close
dbOrcamento.Close

CadastroEspecial_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroEspecial = Nothing
    
    Exit Function
CadastroEspecial_err:
    CadastroEspecial = False
    MsgBox Err.Description
    Resume CadastroEspecial_Fim
End Function

Public Function CadastroOrcamento(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroOrcamento = True
On Error GoTo CadastroOrcamento_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroOrcamento As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroOrcamento = dbOrcamento.QueryDefs("CadastroOrcamento")

With qdfCadastroOrcamento

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    .Parameters("NOME_CLIENTE") = Range("C4").Value
    .Parameters("NOME_CONTATO") = Range("C5").Value
    .Parameters("PROJETO_RTM") = Range("B7").Value
    .Parameters("DTA_PEDIDO") = Range("G3").Value
    .Parameters("DTA_ENTREGA") = Range("G4").Value
    .Parameters("DES_PRODUTO") = Range("G5").Value
    .Parameters("DES_LICENCIADO") = Range("G6").Value
    .Parameters("NOTAFISCAL") = Range("J6").Value
    .Parameters("NF_FATURA") = Range("J7").Value
    
    'QUANTIDADE
    l = 9
    c = 3
    For x = 1 To 8
        .Parameters(x & "QTD") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'FORMATO-1
    .Parameters("1FORMATO") = Cells(11, 3).Value
    .Parameters("2FORMATO") = Cells(11, 7).Value
    
    'FORMATO-2
    .Parameters("3FORMATO") = Cells(12, 3).Value
    .Parameters("4FORMATO") = Cells(12, 7).Value
    
    
    'DESCRIÇÃO
    l = 15
    c = 2
    For x = 1 To 4
        .Parameters(x & "DESCRICAO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'N° DE PÁGINAS
    l = 15
    c = 3
    For x = 1 To 4
        .Parameters(x & "NPAGINAS") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'CORES
    l = 15
    c = 5
    For x = 1 To 4
        .Parameters(x & "CORES") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'PAPEL
    l = 15
    c = 7
    For x = 1 To 4
        .Parameters(x & "PAPEL") = Cells(l, c).Value
        l = l + 1
    Next x
    
    .Parameters("1ACABAMENTO") = Range("C19").Value
    .Parameters("2ACABAMENTO") = Range("C20").Value
    
    .Parameters("OBS") = Range("B22").Value
    .Parameters("DES_ARTIGO") = Range("C26").Value
    .Parameters("DES_PROJETO") = Range("C28").Value
        
    'MÉDICO (1)
    l = 32
    c = 2
    For x = 1 To 3
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MEDICO DIREITO (1)
    l = 32
    c = 4
    For x = 4 To 6
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MÉDICO (2)
    l = 32
    c = 5
    For x = 7 To 9
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MEDICO DIREITO (2)
    l = 32
    c = 8
    For x = 10 To 12
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'FECHADO
    l = 36
    c = 3
    For x = 1 To 8
        .Parameters(x & "FECHADO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'VALOR DA VENDA
    l = 38
    c = 3
    For x = 1 To 8
        .Parameters(x & "VAL_VENDA") = Cells(l, c).Value
        c = c + 1
    Next x

    .Execute
    
End With


'DesbloqueioDeGuia SenhaBloqueio
'
'MarcaTexto IntervaloOrcamento
'MarcaTexto IntervaloEspecial
'
'BloqueioDeGuia SenhaBloqueio

qdfCadastroOrcamento.Close
dbOrcamento.Close



CadastroOrcamento_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroOrcamento = Nothing
    
    Exit Function
CadastroOrcamento_err:
    CadastroOrcamento = False
    MsgBox Err.Description
    Resume CadastroOrcamento_Fim
End Function

Public Function CadastroVenda(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroVenda = True
On Error GoTo CadastroVenda_err
Dim dbOrcamento As DAO.Database
Dim qdfAtualizarVenda As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfAtualizarVenda = dbOrcamento.QueryDefs("CadastroVenda")

With qdfAtualizarVenda
    
    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    'QUANTIDADE
    l = 9
    c = 3
    For x = 1 To 8
        .Parameters(x & "QTD") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'MÉDICO (1)
    l = 32
    c = 2
    For x = 1 To 3
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MEDICO DIREITO (1)
    l = 32
    c = 4
    For x = 4 To 6
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MÉDICO (2)
    l = 32
    c = 5
    For x = 7 To 9
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'MEDICO DIREITO (2)
    l = 32
    c = 8
    For x = 10 To 12
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        l = l + 1
    Next x
    
    'FECHADO
    l = 36
    c = 3
    For x = 1 To 8
        .Parameters(x & "FECHADO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'VALOR DA VENDA
    l = 38
    c = 3
    For x = 1 To 8
        .Parameters(x & "VAL_VENDA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    .Execute
    
End With

qdfAtualizarVenda.Close
dbOrcamento.Close

'admAtualizarDepartamento Range("B1"), Range("J3").Value, strVendedor, "PREVISAO"


CadastroVenda_Fim:

    Set dbOrcamento = Nothing
    Set qdfAtualizarVenda = Nothing
    
    Exit Function
CadastroVenda_err:
    CadastroVenda = False
    MsgBox Err.Description
    Resume CadastroVenda_Fim
End Function

Public Function CarregarOrcamento(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarOrcamento = True
On Error GoTo CarregarOrcamento_err
Dim dbOrcamento As DAO.Database
Dim rstOrcamentoCarregar As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "SELECT Orcamentos.* " & _
         " FROM Orcamentos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstOrcamentoCarregar = dbOrcamento.OpenRecordset(strSQL)

With rstOrcamentoCarregar

    Range("C3").Value = .Fields("VENDEDOR")
    Range("C4").Value = .Fields("CLIENTE")
    Range("C5").Value = .Fields("CONTATO")
    Range("B7").Value = .Fields("PROJETO_PLANILHA_RTM")
    Range("G3").Value = .Fields("DT_PEDIDO")
    Range("G4").Value = .Fields("PREV_ENTREGA")
    Range("G5").Value = .Fields("PRODUTO")
    Range("G6").Value = .Fields("LICENCIADO")
    Range("J3").Value = .Fields("CONTROLE")
    Range("J6").Value = .Fields("NOTA_FISCAL")
    Range("J7").Value = .Fields("NF_FATURA_N")
    
    'QUANTIDADE
    l = 9
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_QUANTIDADE")
        c = c + 1
    Next x
    
    'FORMATO-1
    Cells(11, 3).Value = .Fields("1_FORMATO")
    Cells(11, 7).Value = .Fields("2_FORMATO")
    
    'FORMATO-2
    Cells(12, 3).Value = .Fields("3_FORMATO")
    Cells(12, 7).Value = .Fields("4_FORMATO")
        
    'DESCRIÇÃO
    l = 15
    c = 2
    For x = 1 To 4
        Cells(l, c).Value = .Fields(x & "_DESCRICAO")
        l = l + 1
    Next x
    
    
    'N° DE PÁGINAS
    l = 15
    c = 3
    For x = 1 To 4
        Cells(l, c).Value = .Fields(x & "_N_PAGINAS")
        l = l + 1
    Next x
    
    
    'CORES
    l = 15
    c = 5
    For x = 1 To 4
        Cells(l, c).Value = .Fields(x & "_CORES")
        l = l + 1
    Next x
    
    
    'PAPEL
    l = 15
    c = 7
    For x = 1 To 4
        Cells(l, c).Value = .Fields(x & "_PAPEL")
        l = l + 1
    Next x
    
    Range("C19").Value = .Fields("1_ACABAMENTO")
    Range("C20").Value = .Fields("2_ACABAMENTO")
    
    Range("B22").Value = .Fields("OBSERVACOES")
    Range("C26").Value = .Fields("ARTIGO")
    Range("C28").Value = .Fields("PROJETO")
        
    'MÉDICO (1)
    l = 32
    c = 2
    For x = 1 To 3
        Cells(l, c).Value = .Fields(x & "_MEDICO")
        l = l + 1
    Next x

    'MEDICO DIREITO (1)
    l = 32
    c = 4
    For x = 4 To 6
        Cells(l, c).Value = .Fields(x & "_MEDICO_DIREITO")
        l = l + 1
    Next x

    'MÉDICO (2)
    l = 32
    c = 5
    For x = 7 To 9
        Cells(l, c).Value = .Fields(x & "_MEDICO")
        l = l + 1
    Next x

    'MEDICO DIREITO (2)
    l = 32
    c = 8
    For x = 10 To 12
        Cells(l, c).Value = .Fields(x & "_MEDICO_DIREITO")
        l = l + 1
    Next x
    
    'FECHADO
    l = 36
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_FECHADO")
        c = c + 1
    Next x

    'VALOR DA VENDA
    l = 38
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_VALOR_DA_VENDA")
        c = c + 1
    Next x
    
    'LIBERAÇÃO
    l = 86
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_AUTORIZACAO")
        c = c + 1
    Next x
    
    
    admControleDeIntervalosDeEdicao BaseDeDados, strControle, strVendedor
    
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx [ INTERVALOS DE EDIÇÃO ] xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
'    'ORCAMENTO
'    If .Fields("Orcamento") = True Then
'        IntervaloEditacaoCriar "Orcamento", IntervaloOrcamento
'        DesmarcaTexto IntervaloOrcamento
'    Else
'        MarcaTexto IntervaloOrcamento
'    End If
'
'    'ESPECIAL
'    If .Fields("Especial") = True Then
'        IntervaloEditacaoCriar "Especial", IntervaloEspecial
'        DesmarcaTexto IntervaloEspecial
'    Else
'        MarcaTexto IntervaloEspecial
'    End If
'
'    'VENDA
'    If .Fields("Venda") = True Then
'        IntervaloEditacaoCriar "Venda", IntervaloVenda
'        DesmarcaTexto IntervaloVenda
'    Else
'        MarcaTexto IntervaloVenda
'    End If
'
'    'PREVISÃO
'    If .Fields("Previsao") = True Then
'        IntervaloEditacaoCriar "Previsao", IntervaloPrevisao
'        DesmarcaTexto IntervaloPrevisao
'    Else
'        MarcaTexto IntervaloPrevisao
'    End If
'
'    'RENDIMENTO
'    If .Fields("Rendimento") = True Then
'        IntervaloEditacaoCriar "Rendimento", IntervaloRendimento
'        DesmarcaTexto IntervaloRendimento
'    Else
'        MarcaTexto IntervaloRendimento
'    End If
'
'    'LIBERAÇÃO
'    If .Fields("Liberacao") = True Then
'        IntervaloEditacaoCriar "Liberacao", IntervaloLiberacao
'        DesmarcaTexto IntervaloLiberacao
'    Else
'        MarcaTexto IntervaloLiberacao
'    End If
'
'    'FINANCEIRO
'    If .Fields("Financeiro") = True Then
'        IntervaloEditacaoCriar "Financeiro", IntervaloLiberacao
'        DesmarcaTexto IntervaloLiberacao
'    Else
'        MarcaTexto IntervaloLiberacao
'    End If
        
        
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        
    
End With


CarregarOrcamento_Fim:
    rstOrcamentoCarregar.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstOrcamentoCarregar = Nothing
    
    Exit Function
CarregarOrcamento_err:
    CarregarOrcamento = False
    MsgBox Err.Description
    Resume CarregarOrcamento_Fim
End Function

