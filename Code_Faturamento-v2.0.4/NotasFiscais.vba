Option Compare Database
Option Explicit

Private Sub cmdCopia_Click()
On Error GoTo Err_cmdCopia_Click

Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Dim qry As DAO.Recordset
Dim strPedido As String
Dim resposta As Variant


Set DB = CurrentDb()
'Abre a tabela, adiciona o novo item e atualiza a combo
Set rst = DB.OpenRecordset("NotasFiscais")
Set rst2 = DB.OpenRecordset("NotasFiscaisItens")
Set qry = DB.OpenRecordset("select * FROM NotasFiscaisItens WHERE NotasFiscaisItens.codNotaFiscal = " & Me.Codigo)

resposta = MsgBox("Deseja gerar cópia desta Nota Fiscal?", vbYesNo + vbQuestion, "Cópia de Nota Fiscal")
If resposta = vbYes Then
    With rst
        .AddNew
        
        strPedido = NovoCodigo("NotasFiscais", "codNotaFiscal")
        
        !codNotaFiscal = strPedido
        !codOperacao = Me.codOperacao
        !codCFOP = Me.codCFOP
        !NaturezaDeOperacao = Me.NaturezaDeOperacao
        !DataDeEmissao = Format(Now(), "dd/mm/yy")
        !DataDeSaida = Format(Now(), "dd/mm/yy")
        !codCliente = Me.codCliente
        !Cliente = Me.Cliente
        !Endereco = Me.Endereco
        !Bairro = Me.Bairro
        !Cep = Me.Cep
        !Municipio = Me.Municipio
        !Estado = Me.Estado
        !TelefoneCliente = Me.TelefoneCliente
        !CNPJ = Me.CNPJ
        !IE = Me.IE
        
        !Transportadora = Me.Transportadora
        !Transp_Endereco = Me.Transp_Endereco
        !Transp_Municipio = Me.Transp_Municipio
        !Transp_Estado = Me.Transp_Estado
        !Transp_CNPJ = Me.Transp_CNPJ
        !Transp_IE = Me.Transp_IE
        !FretePorConta = Me.FretePorConta
        !PlacaDoVeiculo = Me.PlacaDoVeiculo
        !UFDaPlaca = Me.UFDaPlaca
        !Quantidade = Me.Quantidade
        !Especie = Me.Especie
        !Marca = Me.Marca
        !NumeroDeControle = Me.NumeroDeControle
        !PesoBruto = Me.PesoBruto
        !PesoLiquido = Me.PesoLiquido
        !DadoAdicional_1 = Me.DadoAdicional_1
        !DadoAdicional_2 = Me.DadoAdicional_2
        !DadoAdicional_3 = Me.DadoAdicional_3
        !DadoAdicional_4 = Me.DadoAdicional_4
        !DadoAdicional_5 = Me.DadoAdicional_5
        !DadoAdicional_6 = Me.DadoAdicional_6
        !DadoAdicional_7 = Me.DadoAdicional_7
        
        !BaseDeCalculoDoICMS = Me.BaseDeCalculoDoICMS
        !ValorDoICMS = Me.ValorDoICMS
        !BaseDeCalculoICMSSubstituicao = Me.BaseDeCalculoICMSSubstituicao
        !ValorDoICMSSubstituicao = Me.ValorDoICMSSubstituicao
        !ValorTotalDosProdutos = Me.ValorTotalDosProdutos
        !ValorDoFrete = Me.ValorDoFrete
        !ValorDoSeguro = Me.ValorDoSeguro
        !OutrasDespesasAcessorias = Me.OutrasDespesasAcessorias
        !ValorTotalDoIPI = Me.ValorTotalDoIPI
        !ValorTotalDaNota = Me.ValorTotalDaNota
        
        
        .Update
        .Close
    End With
        
    With rst2
        
        While Not qry.EOF
            .AddNew
            
            !codNotaFiscal = strPedido
            !DescricaoDoProduto = qry.Fields("DescricaoDoProduto")
            !CST = qry.Fields("CST")
            !Unidade = qry.Fields("Unidade")
            !Quantidade = qry.Fields("Quantidade")
            !ValorUnitario = qry.Fields("ValorUnitario")
            !ValorTotal = qry.Fields("ValorTotal")
            !ICMS = qry.Fields("ICMS")
            
            .Update
            qry.MoveNext
        Wend
        
        .Close

    End With
    Form_Pesquisar.lstCadastro.Requery
    MsgBox "Cópia gerada com sucesso!", vbOKOnly + vbInformation, "Cópia de Nota Fiscal"

    Me.Filter = "codNotaFiscal = " & strPedido
    Me.FilterOn = True
    
    
End If

Exit_cmdCopia_Click:
Set rst = Nothing
Set rst2 = Nothing
Set qry = Nothing
Set DB = Nothing
Exit Sub

Err_cmdCopia_Click:
MsgBox Err.Description
Resume Exit_cmdCopia_Click
End Sub

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    'Visualizar Documento
    stDocName = "NF"
    DoCmd.OpenReport stDocName, acPreview, , "codNotaFiscal = " & Me.Codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
End Sub



Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
       Me.txtDataDeEmissao = Format(Now(), "dd/mm/yy")
       Me.txtDataDeSaida = Format(Now(), "dd/mm/yy")
    End If
    
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub NotasFiscaisItens_Exit(Cancel As Integer)
    Me.ValorTotalDosProdutos = Me.txtSomaProdutos
    Me.ValorTotalDaNota = Me.txtSomaProdutos
    'Me.BaseDeCalculoDoICMS = Me.txtSomaProdutos
End Sub
Private Sub cboCodCFOP_Click()
    Me.txtNaturezaDeOperacao = Me.cbocodCFOP.Column(1)
    Me.qdrcodOperacao = Me.cbocodCFOP.Column(2)
End Sub

Private Sub cboCliente_Click()

    Me.txtCNPJ = ""
    Me.txtEndereco = ""
    Me.txtBairro = ""
    Me.txtCep = ""
    Me.txtMunicipio = ""
    Me.txtTelefoneCliente = ""
    Me.txtEstado = ""
    Me.txtIE = ""


    Me.txtCNPJ = IIf(IsNull(Me.cboCliente.Column(1)), "", Me.cboCliente.Column(1))
    Me.txtEndereco = IIf(IsNull(Me.cboCliente.Column(2)), "", Me.cboCliente.Column(2))
    Me.txtBairro = IIf(IsNull(Me.cboCliente.Column(3)), "", Me.cboCliente.Column(3))
    Me.txtCep = IIf(IsNull(Me.cboCliente.Column(4)), "", Me.cboCliente.Column(4))
    Me.txtMunicipio = IIf(IsNull(Me.cboCliente.Column(5)), "", Me.cboCliente.Column(5))
    Me.txtTelefoneCliente = IIf(IsNull(Me.cboCliente.Column(6)), "", Me.cboCliente.Column(6))
    Me.txtEstado = IIf(IsNull(Me.cboCliente.Column(7)), "", Me.cboCliente.Column(7))
    Me.txtIE = IIf(IsNull(Me.cboCliente.Column(8)), "", Me.cboCliente.Column(8))

End Sub

Private Sub cboTransportadora_Click()

    Me.txtTransp_Endereco = ""
    Me.txtTransp_Municipio = ""
    Me.txtTransp_Estado = ""
    Me.txtTransp_CNPJ = ""
    Me.txtTransp_IE = ""

    Me.txtTransp_Endereco = Me.cboTransportadora.Column(1)
    Me.txtTransp_Municipio = Me.cboTransportadora.Column(2)
    Me.txtTransp_Estado = Me.cboTransportadora.Column(3)
    Me.txtTransp_CNPJ = Me.cboTransportadora.Column(4)
    Me.txtTransp_IE = Me.cboTransportadora.Column(5)

End Sub

Private Sub cboDadoAdicional_1_Click()

Me.txtDadoAdicional_2 = ""
Me.txtDadoAdicional_3 = ""
Me.txtDadoAdicional_4 = ""
Me.txtDadoAdicional_5 = ""
Me.txtDadoAdicional_6 = ""
Me.txtDadoAdicional_7 = ""

Me.txtDadoAdicional_2 = Me.cboDadoAdicional_1.Column(1)
Me.txtDadoAdicional_3 = Me.cboDadoAdicional_1.Column(2)
Me.txtDadoAdicional_4 = Me.cboDadoAdicional_1.Column(3)
Me.txtDadoAdicional_5 = Me.cboDadoAdicional_1.Column(4)
Me.txtDadoAdicional_6 = Me.cboDadoAdicional_1.Column(5)
Me.txtDadoAdicional_7 = Me.cboDadoAdicional_1.Column(6)

End Sub
