Attribute VB_Name = "mod_03_Liberacao"
Option Base 1
Option Explicit

Public Function CadastroLiberacao(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: CadastroLiberacao = True
On Error GoTo CadastroLiberacao_err
Dim dbOrcamento As DAO.Database
Dim qdfadmLiberarOrcamento As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmLiberarOrcamento = dbOrcamento.QueryDefs("CadastroLiberacao")

With qdfadmLiberarOrcamento
    
    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    'AUTORIZACAO
    l = 86
    c = 3
    For x = 1 To 8
        .Parameters(x & "AUTORIZACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    .Execute
    
End With

qdfadmLiberarOrcamento.Close
dbOrcamento.Close

CadastroLiberacao_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmLiberarOrcamento = Nothing
    
    Exit Function
CadastroLiberacao_err:
    CadastroLiberacao = False
    MsgBox Err.Description
    Resume CadastroLiberacao_Fim
End Function




