Attribute VB_Name = "mod_CONFIGURACOES"
'CONTROLES DO SISTEMA
Public Const SenhaBloqueio As String = "RTMBRASIL"

Public Function admOrcamentoEtapaVoltar(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: admOrcamentoEtapaVoltar = True
On Error GoTo admOrcamentoEtapaVoltar_err
Dim dbOrcamento As DAO.Database
Dim qdfadmOrcamentoEtapaVoltar As DAO.QueryDef
Dim strSQL As String


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmOrcamentoEtapaVoltar = dbOrcamento.QueryDefs("admOrcamentoEtapaVoltar")

With qdfadmOrcamentoEtapaVoltar
    
    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    
    .Execute
    
End With

qdfadmOrcamentoEtapaVoltar.Close
dbOrcamento.Close

admOrcamentoEtapaVoltar_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoEtapaVoltar = Nothing
    
    Exit Function
admOrcamentoEtapaVoltar_err:
    admOrcamentoEtapaVoltar = False
    MsgBox Err.Description
    Resume admOrcamentoEtapaVoltar_Fim
End Function

Public Function admOrcamentoEtapaAvancar(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: admOrcamentoEtapaAvancar = True
On Error GoTo admOrcamentoEtapaAvancar_err
Dim dbOrcamento As DAO.Database
Dim qdfadmOrcamentoEtapaAvancar As DAO.QueryDef
Dim strSQL As String


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmOrcamentoEtapaAvancar = dbOrcamento.QueryDefs("admOrcamentoEtapaAvancar")

With qdfadmOrcamentoEtapaAvancar
    
    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    
    .Execute
    
End With

qdfadmOrcamentoEtapaAvancar.Close
dbOrcamento.Close

admOrcamentoEtapaAvancar_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoEtapaAvancar = Nothing
    
    Exit Function
admOrcamentoEtapaAvancar_err:
    admOrcamentoEtapaAvancar = False
    MsgBox Err.Description
    Resume admOrcamentoEtapaAvancar_Fim
End Function


Public Function admControleDeIntervalosDeEdicao(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: admControleDeIntervalosDeEdicao = True
On Error GoTo admControleDeIntervalosDeEdicao_err
Dim dbOrcamento As DAO.Database
Dim rstOrcamento As DAO.Recordset
Dim rstIntervalos As DAO.Recordset
Dim strOrcamento As String
Dim strIntervalos As String


strOrcamento = "SELECT Orcamentos.* " & _
         " FROM Orcamentos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstOrcamento = dbOrcamento.OpenRecordset(strOrcamento)

strIntervalos = "Select * from qryEtapasIntervalosEdicoes where Departamento = '" & rstOrcamento.Fields("Departamento") & "' and Status = '" & rstOrcamento.Fields("Status") & "'"

Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

While Not rstIntervalos.EOF
    
    IntervaloEditacaoCriar rstIntervalos.Fields("Intervalo"), rstIntervalos.Fields("Selecao")
    rstIntervalos.MoveNext

Wend



admControleDeIntervalosDeEdicao_Fim:
    rstOrcamento.Close
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    Set rstOrcamento = Nothing
    
    Exit Function
admControleDeIntervalosDeEdicao_err:
    admControleDeIntervalosDeEdicao = False
    MsgBox Err.Description
    Resume admControleDeIntervalosDeEdicao_Fim


End Function


Public Function MarcarSelecao(BaseDeDados As String)
On Error GoTo BloqueioDeSelecao_err
Dim dbOrcamento As DAO.Database
Dim rstIntervalos As DAO.Recordset
Dim strIntervalos As String

strIntervalos = "qryEtapasIntervalosEdicoes"

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

While Not rstIntervalos.EOF
    
    MarcaTexto rstIntervalos.Fields("Selecao")
    rstIntervalos.MoveNext

Wend

Range("c3").Select

BloqueioDeSelecao_Fim:
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    
    Exit Function
BloqueioDeSelecao_err:
    MsgBox Err.Description
    Resume BloqueioDeSelecao_Fim
End Function


Public Sub LimparSelecao(BaseDeDados As String)
On Error GoTo LimparSelecao_err
Dim dbOrcamento As DAO.Database
Dim rstIntervalos As DAO.Recordset
Dim strIntervalos As String

strIntervalos = "qryEtapasIntervalosEdicoes"

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

While Not rstIntervalos.EOF
    
    LimparTemplate rstIntervalos.Fields("Selecao"), rstIntervalos.Fields("ValorPadrao")
    rstIntervalos.MoveNext

Wend

Range("c3").Select

LimparSelecao_Fim:
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    
    Exit Sub
LimparSelecao_err:
    MsgBox Err.Description
    Resume LimparSelecao_Fim
End Sub

Public Sub admOrcamentoLimpar()
        
    Application.ScreenUpdating = False
    
    IntervaloEditacaoRemoverTodos
        
    LimparSelecao Range("B1")
    
    LimparTemplate "O53:P64", ""
    LimparTemplate "C52", ""

        
    Range("c3").Select
        
    MarcarSelecao Range("B1")
        
    Application.ScreenUpdating = True

End Sub


'Public Const IntervaloOrcamento As String = "C3:E5,G3:H6,J3,J6,J7,B7,B22,C26,C28"
'Public Const IntervaloEspecial As String = "C9:J9,C11:E12,G11:J12,B15:J18,C19:J20"
'Public Const IntervaloVenda As String = "B32:H34,C36:J36,C38:J38"
'
'Public Const IntervaloPrevisao As String = "C52,C53:J63"
'Public Const IntervaloLiberacao As String = "C86:J86"
'Public Const IntervaloRendimento As String = "L53:L64,O53:P64"
'
'Public Const IntervaloFinanceiro As String = "S53:S63"

''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
''#################### [ AMBIENTE DE TESTES ] ####################
''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx


Sub teste_EnviarEmail()

    EnviarEmail "ailtonzsilva@gmail.com", "Teste de envio"

End Sub

Sub Teste_CadastroDeLista()

    DesbloqueioDeGuia SenhaBloqueio
    
    CadastroDeLista "C5", "F", 106, 241
    
    BloqueioDeGuia SenhaBloqueio

End Sub

''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Public Function CarregarClientes(BaseDeDados As String) As Boolean: CarregarClientes = True
On Error GoTo CarregarClientes_err
Dim dbOrcamento As DAO.Database
Dim rstCarregarClientes As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "qryClientes"


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstCarregarClientes = dbOrcamento.OpenRecordset(strSQL)

With rstCarregarClientes

   'ANALISE
    l = 106
    c = 6
      
    
    DesbloqueioDeGuia SenhaBloqueio
        
    While Not rstCarregarClientes.EOF
        Cells(l, c).Value = .Fields("Cliente")
        l = l + 1
        .MoveNext
    Wend
    
    CadastroDeLista "C4", "F", 106, 106 + (rstCarregarClientes.RecordCount - 1)
    
    BloqueioDeGuia SenhaBloqueio
    
End With

rstCarregarClientes.Close
dbOrcamento.Close


CarregarClientes_Fim:

Set rstCarregarClientes = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarClientes_err:
    CarregarClientes = False
    MsgBox Err.Description
    Resume CarregarClientes_Fim
End Function


Public Function CarregarLicenciados(BaseDeDados As String) As Boolean: CarregarLicenciados = True
On Error GoTo CarregarLicenciados_err
Dim dbOrcamento As DAO.Database
Dim rstCarregarLicenciados As DAO.Recordset
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas


strSQL = "qryLicenciados"


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set rstCarregarLicenciados = dbOrcamento.OpenRecordset(strSQL)

With rstCarregarLicenciados

   'ANALISE
    l = 106
    c = 2
      
    
    DesbloqueioDeGuia SenhaBloqueio
    
    
    While Not rstCarregarLicenciados.EOF
        Cells(l, c).Value = .Fields("Licenciado")
        Cells(l, c + 1).Value = .Fields("Direitos")
        Cells(l, c + 2).Value = .Fields("Margem")
        l = l + 1
        .MoveNext
    Wend
    
    CadastroDeLista "G6", "B", 106, 106 + (rstCarregarLicenciados.RecordCount - 1)
    
    BloqueioDeGuia SenhaBloqueio
    
End With

rstCarregarLicenciados.Close
dbOrcamento.Close


CarregarLicenciados_Fim:

Set rstCarregarLicenciados = Nothing
Set dbOrcamento = Nothing

    
    Exit Function
CarregarLicenciados_err:
    CarregarLicenciados = False
    MsgBox Err.Description
    Resume CarregarLicenciados_Fim
End Function
