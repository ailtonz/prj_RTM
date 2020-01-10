Attribute VB_Name = "mod_02_Producao"
Option Base 1
Option Explicit

Public Function CadastroPrevisao(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroPrevisao = True
On Error GoTo CadastroPrevisao_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroPrevisao As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroPrevisao = dbOrcamento.QueryDefs("CadastroPrevisao")

With qdfCadastroPrevisao

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    .Parameters("RESP_PRODUCAO") = Range("C52").Value
    
    'TRADUÇÃO
    l = 53
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRADUCAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'REVISÃO ORTOGRÁFICA
    l = 54
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVORTOGRAFICA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    
    'REVISÃO MÉDICA
    l = 55
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVMEDICA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    
    'CRIAÇÃO
    l = 56
    c = 3
    For x = 1 To 8
        .Parameters(x & "CRIACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'ILUSTRAÇÃO/ DIAGRAM
    l = 57
    c = 3
    For x = 1 To 8
        .Parameters(x & "ILUSTRACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'DIAGRAMAÇÃO
    l = 58
    c = 3
    For x = 1 To 8
        .Parameters(x & "DIAGRAMACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'PAPEL
    l = 59
    c = 3
    For x = 1 To 8
        .Parameters(x & "PAPEL") = Cells(l, c).Value
        c = c + 1
    Next x
                
    'IMPRESSÃO
    l = 60
    c = 3
    For x = 1 To 8
        .Parameters(x & "IMPRESSAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'PAPEL/IMPRESSÃO
    l = 61
    c = 3
    For x = 1 To 8
        .Parameters(x & "PAPELIMPRESSAO") = Cells(l, c).Value
        c = c + 1
    Next x
                
    'TRANSPORTE
    l = 62
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRANSPORTE") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'OUTROS
    l = 63
    c = 3
    For x = 1 To 8
        .Parameters(x & "OUTROS") = Cells(l, c).Value
        c = c + 1
    Next x
    
    .Execute
    
    
End With

qdfCadastroPrevisao.Close
dbOrcamento.Close


CadastroPrevisao_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroPrevisao = Nothing
    
    Exit Function
CadastroPrevisao_err:
    CadastroPrevisao = False
    MsgBox Err.Description
    Resume CadastroPrevisao_Fim
End Function

Public Function CarregarPrevisao(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarPrevisao = True
On Error GoTo CarregarPrevisao_err
Dim dbOrcamento As DAO.Database
Dim rstPrevisaoCarregar As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "SELECT PrevisoesDeCustos.* " & _
         " FROM PrevisoesDeCustos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstPrevisaoCarregar = dbOrcamento.OpenRecordset(strSQL)

With rstPrevisaoCarregar

    Range("C52").Value = .Fields("RESPONSAVEL_PRODUCAO")
    
    'TRADUÇÃO
    l = 53
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_TRADUCAO")
        c = c + 1
    Next x
    
    'REVISÃO ORTOGRÁFICA
    l = 54
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_REVISAO_ORTOGRAFICA")
        c = c + 1
    Next x
    
    
    'REVISÃO MÉDICA
    l = 55
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_REVISAO_MEDICA")
        c = c + 1
    Next x
    
    
    'CRIAÇÃO
    l = 56
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_CRIACAO")
        c = c + 1
    Next x
    
    'ILUSTRAÇÃO/ DIAGRAM
    l = 57
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_ILUSTRACAO_DIAGRAM")
        c = c + 1
    Next x
    
    'DIAGRAMAÇÃO
    l = 58
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_DIAGRAMACAO")
        c = c + 1
    Next x
    
    'PAPEL
    l = 59
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_PAPEL")
        c = c + 1
    Next x
                
    'IMPRESSÃO
    l = 60
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_IMPRESSAO")
        c = c + 1
    Next x
    
    'PAPEL/IMPRESSÃO
    l = 61
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_PAPEL_IMPRESSAO")
        c = c + 1
    Next x
                
    'TRANSPORTE
    l = 62
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_TRANSPORTE")
        c = c + 1
    Next x
    
    'OUTROS
    l = 63
    c = 3
    For x = 1 To 8
        Cells(l, c).Value = .Fields(x & "_OUTROS")
        c = c + 1
    Next x
    
End With

rstPrevisaoCarregar.Close
dbOrcamento.Close

CarregarPrevisao_Fim:

    Set rstPrevisaoCarregar = Nothing
    Set dbOrcamento = Nothing
    
    Exit Function
CarregarPrevisao_err:
    CarregarPrevisao = False
    MsgBox Err.Description
    Resume CarregarPrevisao_Fim
End Function

