Option Compare Database
Option Explicit

Public strTabela As String

Public Function Posicao(campo As String, retorno As String) As Integer

Dim rPosicao As DAO.Recordset

Set rPosicao = CurrentDb.OpenRecordset("Select * from qryPosicao where campo = '" & campo & "'")

If retorno = "Linha" Then
    Posicao = rPosicao.Fields("Linha")
ElseIf retorno = "Coluna" Then
    Posicao = rPosicao.Fields("Coluna")
End If

rPosicao.Close

End Function

Public Function NovoCodigo(Tabela, campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function Chancelamento(Inicio As Integer, Final As Integer) As String

Dim ch_X As Boolean
Dim Texto As String

ch_X = True

For Inicio = 1 To Final
    Texto = Texto + IIf(ch_X, "x", "-")
    ch_X = Not ch_X
Next

Chancelamento = Texto

End Function

Public Function Comissao(valParcela, valAvista, valComissao)
    Comissao = (valComissao * (valParcela / valAvista * 100)) / 100
End Function

Public Function AbrirArquivo(sTitulo As String, sDecricao As String, sTipo As String, SelecaoMultipla As Boolean) As String
Dim fd As Office.FileDialog

'Di�logo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'T�tulo
fd.Title = sTitulo

'Filtros e descri��o dos mesmos
fd.Filters.Add sDecricao, sTipo

'Premiss�es de sela��o
fd.AllowMultiSelect = SelecaoMultipla

If fd.Show = -1 Then
    AbrirArquivo = fd.SelectedItems(1)
End If

End Function

Public Function ExportarDuplicatas(val As String)
On Error GoTo Err_ExportarDuplicatas

Dim Origem As DAO.Recordset

Set Origem = CurrentDb.OpenRecordset("Select * from Lubriquim_Duplicatas where ImportacaoReceita = 0")

If Origem.RecordCount > 0 Then

    Dim dbDestino As DAO.Database
    Dim Destino As DAO.Recordset
    Dim rstTabela As DAO.Recordset
    
    Dim objAccess As Object
    
    Dim strArq As String
    Dim strSQL As String
    Dim codFaturamento As Long
    Dim cont As Long
    Dim s As Variant
    Dim I As Long
    
    Set objAccess = CreateObject("Access.Application")
    
    strArq = AbrirArquivo("Informe o destino das duplicatas", "BDs do Access", "*.MDB;*.MDE", False)
    
    
    'Se selecionou arquivo, atualiza os v�nculos
    If strArq <> "" Then
        
        objAccess.OpenCurrentDatabase (strArq)
        objAccess.Visible = False
        
        Set dbDestino = objAccess.CurrentDb
        Set Destino = dbDestino.OpenRecordset("Select * from Faturamentos")
            
        Set rstTabela = dbDestino.OpenRecordset("SELECT Max([codFaturamento])+1 AS CodigoNovo FROM Faturamentos;")
        
        codFaturamento = IIf(IsNull(rstTabela.Fields("CodigoNovo")), 1, rstTabela.Fields("CodigoNovo"))
            
        Origem.MoveLast
        cont = Origem.RecordCount
        Origem.MoveFirst
        
        BeginTrans
        
        s = SysCmd(acSysCmdInitMeter, "Exportando " & cont & " Duplicatas", cont)
        
        I = 0
        
        DoCmd.Hourglass True
        
        While Not Origem.EOF
                            
            Destino.AddNew
            Destino.Fields("codFaturamento") = codFaturamento
            Destino.Fields("codNotaFiscal") = Origem.Fields("codNotaFiscal")
            Destino.Fields("Razao") = Origem.Fields("Apelido")
            Destino.Fields("DataDeEmissao") = Origem.Fields("Emissao")
            Destino.Fields("DataDeVencimento") = Origem.Fields("Vencimento")
            Destino.Fields("ValorDoFaturamento") = Origem.Fields("Valor")
            Destino.Fields("DescricaoDoFaturamento") = Origem.Fields("Ordem")
            Destino.Fields("codTipoFaturamento") = "Faturamento"
            Destino.Fields("Status") = "Aberto"
            Destino.Update
            
            codFaturamento = codFaturamento + 1
            
            Origem.Edit
            Origem.Fields("ImportacaoReceita") = -1
            Origem.Update
            
            Origem.MoveNext
            
            I = I + 1
            
            s = SysCmd(acSysCmdUpdateMeter, I)
            
        Wend
        
        CommitTrans
        
        objAccess.CloseCurrentDatabase
    
        objAccess.Quit
        
        Set objAccess = Nothing
        Set Origem = Nothing
        Set Destino = Nothing
        Set dbDestino = Nothing
        Set rstTabela = Nothing
        
        DoCmd.Hourglass False
        ' Hide the Meter
        s = SysCmd(acSysCmdSetStatus, " ")
        
        MsgBox "Duplicatas enviadas com sucesso.", vbOKOnly + vbInformation, "Exporta��o de Duplicatas ao financeiro"
    
    End If

Else
    Set Origem = Nothing
    DoCmd.Hourglass False
    MsgBox "N�o h� Duplicatas a serem enviadas ao financeiro.", vbOKOnly + vbExclamation, "Exporta��o de Duplicatas ao financeiro"
    Exit Function
End If

Exit_ExportarDuplicatas:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarDuplicatas:
    MsgBox Err.Description
    Resume Exit_ExportarDuplicatas

End Function


Public Function ExportarComissoes()
    On Error GoTo Err_ExportarComissoes

Dim Origem As DAO.Recordset
Dim strOrigem As String

Dim Duplicatas As DAO.Recordset
Dim strDuplicatas As String

strOrigem = "SELECT " & _
            "   qryComissao.codPedidoLubriquim, qryComissao.codNotaFiscal, " & _
            "   qryComissao.Funcionario, Sum(qryComissao.ValorComissao) AS ValComissao, qryComissao.ValorTotalDaNota " & _
            "FROM " & _
            "   (SELECT " & _
            "   Lubriquim_PedidosVendasComissoes.Funcionario, Lubriquim_PedidosVendasComissoes.codPedidoLubriquim, " & _
            "   Lubriquim_NotasFiscais.codNotaFiscal, Lubriquim_PedidosVendasComissoes.ValorComissao, " & _
            "   Lubriquim_NotasFiscais.ValorTotalDaNota, Lubriquim_PedidosVendasComissoes.PG " & _
            "FROM " & _
            "   (Lubriquim_PedidosVendas INNER JOIN Lubriquim_PedidosVendasComissoes ON Lubriquim_PedidosVendas.codPedidoLubriquim = Lubriquim_PedidosVendasComissoes.codPedidoLubriquim) " & _
            "   INNER JOIN Lubriquim_NotasFiscais ON Lubriquim_PedidosVendas.codPedidoLubriquim = Lubriquim_NotasFiscais.codPedidoLubriquim " & _
            "WHERE Lubriquim_PedidosVendasComissoes.PG = No) as qryComissao " & _
            "GROUP BY " & _
            "   qryComissao.codPedidoLubriquim, qryComissao.codNotaFiscal, " & _
            "   qryComissao.Funcionario, qryComissao.ValorTotalDaNota " & _
            "ORDER BY " & _
            "   qryComissao.codNotaFiscal"


Set Origem = CurrentDb.OpenRecordset(strOrigem)

If Origem.RecordCount > 0 Then
    DoCmd.Hourglass True
    While Not Origem.EOF
        strDuplicatas = "SELECT  " & _
                        "   Lubriquim_Duplicatas.codDuplicata, Lubriquim_Duplicatas.Apelido,  " & _
                        "   Lubriquim_Duplicatas.codNotaFiscal, Lubriquim_Duplicatas.Ordem, Lubriquim_Duplicatas.Valor " & _
                        "FROM  " & _
                        "   Lubriquim_Duplicatas " & _
                        "WHERE  " & _
                        "   (((Lubriquim_Duplicatas.codNotaFiscal)=" & Origem.Fields("codNotaFiscal") & ")) " & _
                        "ORDER BY  " & _
                        "   Lubriquim_Duplicatas.codDuplicata "
        Set Duplicatas = CurrentDb.OpenRecordset(strDuplicatas)
        If Duplicatas.RecordCount > 0 Then
            While Not Duplicatas.EOF
                MsgBox Duplicatas.Fields("codNotaFiscal")
                MsgBox Comissao(Duplicatas.Fields("Valor"), Origem.Fields("ValorTotalDaNota"), Origem.Fields("ValComissao"))
                Duplicatas.MoveNext
            Wend
        End If
        Origem.MoveNext
    Wend

End If

Set Origem = Nothing
Set Duplicatas = Nothing


Exit_ExportarComissoes:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarComissoes:
    MsgBox Err.Description
    Resume Exit_ExportarComissoes
End Function

Public Function ExportarMateriaPrima()
On Error GoTo Err_ExportarMateriaPrima

Dim Origem As DAO.Recordset

Set Origem = CurrentDb.OpenRecordset("Select * from FAT_ProdutosMateriasPrimas")

If Origem.RecordCount > 0 Then

    Dim dbDestino As DAO.Database
    Dim Destino As DAO.Recordset
    Dim rstTabela As DAO.Recordset
    
    Dim objAccess As Object
    
    Dim strArq As String
    Dim strSQL As String
    Dim codFaturamento As Long
    Dim cont As Long
    Dim s As Variant
    Dim I As Long
    
    Set objAccess = CreateObject("Access.Application")
    
    strArq = AbrirArquivo("Informe o destino", "BDs do Access", "*.MDB;*.MDE", False)
    
    
    'Se selecionou arquivo, atualiza os v�nculos
    If strArq <> "" Then
        
        objAccess.OpenCurrentDatabase (strArq)
        objAccess.Visible = False
        
        Set dbDestino = objAccess.CurrentDb
        Set Destino = dbDestino.OpenRecordset("Select * from FAT_Produtos")
            
        Set rstTabela = dbDestino.OpenRecordset("SELECT Max([codProduto])+1 AS CodigoNovo FROM FAT_Produtos;")
        
        codFaturamento = IIf(IsNull(rstTabela.Fields("CodigoNovo")), 1, rstTabela.Fields("CodigoNovo"))
            
        Origem.MoveLast
        cont = Origem.RecordCount
        Origem.MoveFirst
        
        BeginTrans
        
        s = SysCmd(acSysCmdInitMeter, "Exportando " & cont, cont)
        
        I = 0
        
        DoCmd.Hourglass True
        
        While Not Origem.EOF
                            
            Destino.AddNew
            Destino.Fields("codProduto") = codFaturamento
            Destino.Fields("DescricaoDoProduto") = Origem.Fields("MateriaPrima")
            Destino.Fields("codInterno") = Origem.Fields("CodigoInterno")
            Destino.Fields("codTipo") = "MP"
            Destino.Fields("Valor") = Origem.Fields("Valor")
            Destino.Update
            
            codFaturamento = codFaturamento + 1
                        
            Origem.MoveNext
            
            I = I + 1
            
            s = SysCmd(acSysCmdUpdateMeter, I)
            
        Wend
        
        CommitTrans
        
        objAccess.CloseCurrentDatabase
    
        objAccess.Quit
        
        Set objAccess = Nothing
        Set Origem = Nothing
        Set Destino = Nothing
        Set dbDestino = Nothing
        Set rstTabela = Nothing
        
        DoCmd.Hourglass False
        ' Hide the Meter
        s = SysCmd(acSysCmdSetStatus, " ")
        
        MsgBox "Enviadas com sucesso.", vbOKOnly + vbInformation, "Exporta��o"
    
    End If

Else
    Set Origem = Nothing
    DoCmd.Hourglass False
    MsgBox "N�o h� itens a serem enviados.", vbOKOnly + vbExclamation, "Exporta��o"
    Exit Function
End If

Exit_ExportarMateriaPrima:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarMateriaPrima:
    MsgBox Err.Description
    Resume Exit_ExportarMateriaPrima

End Function

Public Function Confirmar(sMensagem As String) As _
Boolean
'Faz uma pergunta ao usu�rio e retorma True se a
'resposta for SIM, e false se a resposta for N�O
Dim intResp As Integer

intResp = MsgBox(sMensagem, vbYesNo + vbQuestion, _
"Confirma��o")

If intResp = vbYes Then
    Confirmar = True
Else
    Confirmar = False
End If
End Function

