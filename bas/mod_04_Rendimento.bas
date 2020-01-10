Attribute VB_Name = "mod_04_Rendimento"
Option Base 1
Option Explicit

Public Function CadastroRendimento(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroRendimento = True
On Error GoTo CadastroRendimento_err
Dim dbOrcamento As DAO.Database
Dim qdfCadastroRendimento As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfCadastroRendimento = dbOrcamento.QueryDefs("CadastroRendimento")

With qdfCadastroRendimento

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    'REALIZADO
    l = 53
    c = 12
    For x = 1 To 12
        .Parameters(x & "REALIZADO") = Cells(l, c).Value
        l = l + 1
    Next x
   
    
    'FORNECEDOR
    l = 53
    c = 15
    For x = 1 To 12
        .Parameters(x & "FORNECEDOR") = Cells(l, c).Value
        l = l + 1
    Next x
    
    
    'FORNECEDOR NF
    l = 53
    c = 16
    For x = 1 To 12
        .Parameters(x & "FORNECEDORNF") = Cells(l, c).Value
        l = l + 1
    Next x
    
    .Execute
    
    
End With

qdfCadastroRendimento.Close
dbOrcamento.Close


CadastroRendimento_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroRendimento = Nothing
    
    Exit Function
CadastroRendimento_err:
    CadastroRendimento = False
    MsgBox Err.Description
    Resume CadastroRendimento_Fim
End Function

Public Function CarregarRendimento(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CarregarRendimento = True
On Error GoTo CarregarRendimento_err
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


    'REALIZADO
    l = 53
    c = 12
    For x = 1 To 12
        Cells(l, c).Value = .Fields(x & "_REALIZADO")
        l = l + 1
    Next x


    'FORNECEDOR
    l = 53
    c = 15
    For x = 1 To 12
        Cells(l, c).Value = .Fields(x & "_FORNECEDOR")
        l = l + 1
    Next x
    
    
    'FORNECEDOR NF
    l = 53
    c = 16
    For x = 1 To 12
        Cells(l, c).Value = .Fields(x & "_FORNECEDOR_NF")
        l = l + 1
    Next x

    
End With

rstRendimentoCarregar.Close
dbOrcamento.Close


CarregarRendimento_Fim:

Set rstRendimentoCarregar = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarRendimento_err:
    CarregarRendimento = False
    MsgBox Err.Description
    Resume CarregarRendimento_Fim
End Function



