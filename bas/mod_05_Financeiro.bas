Attribute VB_Name = "mod_05_Financeiro"
Option Base 1
Option Explicit

Public Function CadastroFinanceiro(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroFinanceiro = True
On Error GoTo CadastroFinanceiro_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroFinanceiro As DAO.QueryDef
Dim strSQL As String
Dim l As Integer, c As Integer
Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroFinanceiro = dbOrcamento.QueryDefs("CadastroFinanceiro")

With qdfCadastroFinanceiro
    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
   'ANALISE
    l = 53
    c = 19
    For x = 1 To 11
        .Parameters(x & "ANALISE") = Cells(l, c).Value
        l = l + 1
    Next x
    .Execute
End With

qdfCadastroFinanceiro.Close
dbOrcamento.Close

CadastroFinanceiro_Fim:
    Set dbOrcamento = Nothing
    Set qdfCadastroFinanceiro = Nothing
    Exit Function
CadastroFinanceiro_err:
    CadastroFinanceiro = False
    MsgBox Err.Description
    Resume CadastroFinanceiro_Fim
End Function

Public Function CarregarFinanceiro(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarFinanceiro = True
On Error GoTo CarregarFinanceiro_err
Dim dbOrcamento As DAO.Database
Dim rstRendimentoCarregar As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "SELECT PrevisoesDeCustos.* " & _
         " FROM PrevisoesDeCustos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstRendimentoCarregar = dbOrcamento.OpenRecordset(strSQL)

With rstRendimentoCarregar

   'ANALISE
    l = 53
    c = 19
    For x = 1 To 11
        Cells(l, c).Value = .Fields(x & "_ANALISE")
        l = l + 1
    Next x
    
End With

rstRendimentoCarregar.Close
dbOrcamento.Close


CarregarFinanceiro_Fim:

Set rstRendimentoCarregar = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarFinanceiro_err:
    CarregarFinanceiro = False
    MsgBox Err.Description
    Resume CarregarFinanceiro_Fim
End Function

Public Sub copiarProduzido()

Dim LinhaOrigem As Integer: LinhaOrigem = 76
Dim ColunaOrigem As Integer: ColunaOrigem = 11

Dim LinhaDestino As Integer: LinhaDestino = 53
Dim ColunaDestino As Integer: ColunaDestino = 19

Dim x As Integer

DesbloqueioDeGuia SenhaBloqueio

For x = 1 To 10
    Cells(LinhaDestino, ColunaDestino).Value = Cells(LinhaOrigem, ColunaOrigem).Value
    LinhaOrigem = LinhaOrigem + 1
    LinhaDestino = LinhaDestino + 1
Next x

Range("c3").Select
BloqueioDeGuia SenhaBloqueio

End Sub

Public Sub copiarAnalise()

Dim LinhaOrigem As Integer: LinhaOrigem = 53
Dim ColunaOrigem As Integer: ColunaOrigem = 19

Dim LinhaDestino As Integer: LinhaDestino = 76
Dim ColunaDestino As Integer: ColunaDestino = 12

Dim x As Integer

DesbloqueioDeGuia SenhaBloqueio

For x = 1 To 10
    Cells(LinhaDestino, ColunaDestino).Value = Cells(LinhaOrigem, ColunaOrigem).Value
    LinhaOrigem = LinhaOrigem + 1
    LinhaDestino = LinhaDestino + 1
Next x

Range("c3").Select
BloqueioDeGuia SenhaBloqueio

End Sub


Public Function CadastroFinal(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroFinal = True
On Error GoTo CadastroFinal_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroFinal As DAO.QueryDef
Dim strSQL As String
Dim l As Integer, c As Integer
Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroFinal = dbOrcamento.QueryDefs("CadastroFinal")

With qdfCadastroFinal
    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
   'ANALISE
    l = 76
    c = 12
    For x = 1 To 10
        .Parameters(x & "Final") = Cells(l, c).Value
        l = l + 1
    Next x
    .Execute
End With



qdfCadastroFinal.Close
dbOrcamento.Close

CadastroFinal_Fim:
    Set dbOrcamento = Nothing
    Set qdfCadastroFinal = Nothing
    Exit Function
CadastroFinal_err:
    CadastroFinal = False
    MsgBox Err.Description
    Resume CadastroFinal_Fim
End Function

Public Function CarregarFinal(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarFinal = True
On Error GoTo CarregarFinal_err
Dim dbOrcamento As DAO.Database
Dim rstRendimentoCarregar As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "SELECT PrevisoesDeCustos.* " & _
         " FROM PrevisoesDeCustos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstRendimentoCarregar = dbOrcamento.OpenRecordset(strSQL)

With rstRendimentoCarregar

   'ANALISE
    l = 76
    c = 12
    For x = 1 To 10
        Cells(l, c).Value = .Fields(x & "_FINAL")
        l = l + 1
    Next x
    
End With

rstRendimentoCarregar.Close
dbOrcamento.Close


CarregarFinal_Fim:

Set rstRendimentoCarregar = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarFinal_err:
    CarregarFinal = False
    MsgBox Err.Description
    Resume CarregarFinal_Fim
End Function

Public Function CadastroIndice(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroIndice = True
On Error GoTo CadastroIndice_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroIndice As DAO.QueryDef
Dim strSQL As String
Dim l As Integer, c As Integer
Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroIndice = dbOrcamento.QueryDefs("CadastroIndice")

With qdfCadastroIndice
    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
   'ANALISE
    l = 80
    c = 19
    For x = 1 To 10
        .Parameters(x & "Indice") = Cells(l, c).Value
        l = l + 1
    Next x
    .Execute
End With



qdfCadastroIndice.Close
dbOrcamento.Close

CadastroIndice_Fim:
    Set dbOrcamento = Nothing
    Set qdfCadastroIndice = Nothing
    Exit Function
CadastroIndice_err:
    CadastroIndice = False
    MsgBox Err.Description
    Resume CadastroIndice_Fim
End Function

Public Function CarregarIndice(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarIndice = True
On Error GoTo CarregarIndice_err
Dim dbOrcamento As DAO.Database
Dim rstRendimentoCarregar As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "SELECT Orcamentos.* " & _
         " FROM Orcamentos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstRendimentoCarregar = dbOrcamento.OpenRecordset(strSQL)

With rstRendimentoCarregar

   'ANALISE
    l = 80
    c = 19
    For x = 1 To 3
        Cells(l, c).Value = .Fields(x & "_Indice")
        l = l + 1
    Next x
    
End With

rstRendimentoCarregar.Close
dbOrcamento.Close


CarregarIndice_Fim:

Set rstRendimentoCarregar = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarIndice_err:
    CarregarIndice = False
    MsgBox Err.Description
    Resume CarregarIndice_Fim
End Function

